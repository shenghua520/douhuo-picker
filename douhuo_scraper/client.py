"""抖货选品中心 API 客户端：Playwright 持久化登录 + 页面内 fetch。"""

from __future__ import annotations

import json
import random
import time
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, urlparse

from playwright.sync_api import Page, sync_playwright

BASE = "https://www.douhuomall.com"
ENTRY_URL = BASE + "/choice/v3/#/index"
LOGIN_URL = BASE + "/choice/v3/#/register_login"

API = {
    "user_info": "/manager/Choice/getUserInfo",
    "banner_img": "/manager/Choice/getBannerImg",
    "recs_special": "/manager/Choice/getRecsSpecial",
    "top_special": "/manager/choice/getTopSpecial",
    "navi_special": "/manager/choice/getNaviSpecial",
    "goods_list": "/manager/Choice/getGoodsList",
    "top_goods": "/manager/Choice/getTopGoods",
    "hot_goods": "/manager/Choice/getHotGoods",
    "goods_detail": "/manager/Choice/getGoodsDetail",
    "top_cate": "/manager/Choice/getTopCate",
    "supply": "/manager/Choice/getSupply",
}

FETCH_JS = """
async ([url, body, timeoutMs]) => {
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), timeoutMs);
  try {
    const res = await fetch(url, {
      method: 'POST',
      credentials: 'include',
      headers: {'Content-Type': 'application/json;charset=UTF-8'},
      body: JSON.stringify(body),
      signal: ctrl.signal,
    });
    return {ok: res.ok, status: res.status, text: await res.text()};
  } catch (e) {
    return {ok: false, status: 0, text: String((e && e.message) || e)};
  } finally {
    clearTimeout(timer);
  }
}
"""


class ApiError(RuntimeError):
    def __init__(self, code: int, msg: str, url: str):
        super().__init__(f"{url} -> {code} {msg}")
        self.code, self.msg, self.url = code, msg, url


class NotLoggedIn(RuntimeError):
    pass


class DouhuoClient:
    def __init__(
        self,
        profile_dir: Path,
        headless: bool = True,
        timeout: int = 30,
        retries: int = 4,
        min_delay: float = 0.35,
        max_delay: float = 0.8,
        cookies_path: Path | None = None,
    ):
        self.profile_dir = Path(profile_dir)
        self.headless = headless
        self.timeout_ms = timeout * 1000
        self.retries = retries
        self.min_delay = min_delay
        self.max_delay = max_delay
        self.cookies_path = Path(cookies_path) if cookies_path else self.profile_dir / "cookies.json"
        self._pw = None
        self._ctx = None
        self._page: Page | None = None

    def _apply_saved_cookies(self) -> None:
        if not self.cookies_path.exists():
            return
        try:
            cookies = json.loads(self.cookies_path.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            return
        if cookies:
            self._ctx.add_cookies(cookies)

    def save_cookies(self) -> Path:
        self.cookies_path.parent.mkdir(parents=True, exist_ok=True)
        cookies = self._ctx.cookies()
        self.cookies_path.write_text(json.dumps(cookies, ensure_ascii=False, indent=2), encoding="utf-8")
        return self.cookies_path

    def __enter__(self) -> "DouhuoClient":
        self._pw = sync_playwright().start()
        self._ctx = self._pw.chromium.launch_persistent_context(
            user_data_dir=str(self.profile_dir),
            headless=self.headless,
            locale="zh-CN",
            viewport={"width": 1440, "height": 900},
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"
            ),
            args=["--disable-blink-features=AutomationControlled"],
        )
        self._page = self._ctx.pages[0] if self._ctx.pages else self._ctx.new_page()
        self._apply_saved_cookies()
        self.open_entry()
        return self

    def __exit__(self, *exc) -> None:
        try:
            if self._ctx:
                self._ctx.close()
        finally:
            if self._pw:
                self._pw.stop()

    def open_entry(self, url: str = ENTRY_URL) -> None:
        assert self._page is not None
        last = None
        for attempt in range(1, self.retries + 1):
            try:
                self._page.goto(url, wait_until="domcontentloaded", timeout=self.timeout_ms)
                self._page.wait_for_timeout(1800)
                return
            except Exception as exc:  # noqa: BLE001
                last = exc
                time.sleep(attempt * 1.5)
        raise RuntimeError(f"页面加载失败: {last}")

    def post(self, path: str, payload: dict | None = None) -> Any:
        assert self._page is not None
        body = payload or {}
        last: Exception | None = None
        for attempt in range(1, self.retries + 1):
            time.sleep(random.uniform(self.min_delay, self.max_delay))
            try:
                raw = self._page.evaluate(FETCH_JS, [path, body, self.timeout_ms])
            except Exception as exc:  # noqa: BLE001
                last = exc
                time.sleep(attempt * 1.2)
                continue
            if not raw.get("ok"):
                last = RuntimeError(f"HTTP {raw.get('status')} {str(raw.get('text'))[:180]}")
                time.sleep(attempt * 1.2)
                continue
            try:
                data = json.loads(raw["text"])
            except json.JSONDecodeError as exc:
                raise NotLoggedIn(f"{path} 非 JSON，可能未登录") from exc
            code = data.get("error_code")
            msg = data.get("error_msg") or ""
            if code in (1011, 1020):
                raise NotLoggedIn(f"{path}: {code} {msg}")
            if code:
                raise ApiError(int(code), msg, path)
            return data.get("result")
        raise RuntimeError(f"{path} 失败: {last}")

    def user_info(self) -> dict:
        return self.post(API["user_info"], {}) or {}

    def login_state(self) -> str:
        try:
            info = self.user_info()
        except NotLoggedIn:
            return "未登录"
        if str(info.get("is_tourist")) == "1":
            return "游客"
        if str(info.get("sys_id")) == "11":
            return "试用账号"
        return f"已登录:{info.get('sys_name') or info.get('sys_id')}"

    def ensure_login(self, wait_minutes: int = 12) -> str:
        state = self.login_state()
        if state.startswith("已登录"):
            path = self.save_cookies()
            print(f">>> 已是正式账号 {state}，cookies -> {path}")
            return state
        if self.headless:
            raise NotLoggedIn("需要登录。请运行: python douhuo_scraper/run.py --login")
        print("\n>>> 请在弹出浏览器完成登录（扫码/短信）。检测到正式账号后自动继续。")
        try:
            self.open_entry(LOGIN_URL)
        except Exception:  # noqa: BLE001
            pass
        deadline = time.time() + wait_minutes * 60
        last_print = 0.0
        while time.time() < deadline:
            time.sleep(3)
            try:
                self.open_entry()
                state = self.login_state()
            except Exception as exc:  # noqa: BLE001
                print(">>> 检测登录时出错，重试:", exc)
                continue
            if state.startswith("已登录"):
                # 再确认一次，避免瞬时脏读
                time.sleep(2)
                state2 = self.login_state()
                if state2.startswith("已登录"):
                    path = self.save_cookies()
                    print(f">>> 登录成功: {state2}")
                    print(f">>> cookies 已保存: {path}")
                    return state2
            if time.time() - last_print > 20:
                print(f">>> 当前登录态: {state} …请完成登录")
                last_print = time.time()
        raise NotLoggedIn(f"{wait_minutes} 分钟内未检测到正式账号登录成功")

    def banner(self) -> dict:
        return self.post(API["banner_img"], {"type": 1, "navi": 0}) or {}

    def recs_special(self) -> list[dict]:
        return self.post(API["recs_special"], {"site": 1, "type": 3, "navi": 0}) or []

    def top_special(self) -> list[dict]:
        return self.post(API["top_special"], {"site": 1, "type": 3, "navi": 0}) or []

    def navi_special(self) -> list[dict]:
        return self.post(API["navi_special"], {"site": 1, "type": 3, "navi": 0}) or []

    def goods_list(
        self,
        *,
        nav_type: int = 0,
        nav_id: str | int = "",
        page: int = 1,
        limit: int = 100,
        **extra: Any,
    ) -> dict:
        payload = {
            "nav_type": nav_type,
            "nav_id": str(nav_id) if nav_id not in ("", None) else "",
            "type": 0,
            "spu_sn": "",
            "spu_name": "",
            "sort_field": 0,
            "sort_desc": 0,
            "page": page,
            "limit": limit,
            "supply_type": [],
            "cate_id": 0,
            "brand_id": [],
            "tag_id": "",
            "min_profit": 0,
            "max_profit": 0,
            "min_plat": 0,
            "max_plat": 0,
            "is_freight": "",
        }
        payload.update(extra)
        return self.post(API["goods_list"], payload) or {}

    def iter_goods(
        self,
        *,
        nav_type: int = 0,
        nav_id: str | int = "",
        limit: int = 100,
        max_items: int | None = None,
        start_page: int = 1,
    ):
        page = start_page
        got = 0
        total = None
        pages = None
        while True:
            res = self.goods_list(nav_type=nav_type, nav_id=nav_id, page=page, limit=limit)
            total = res.get("total")
            pages = res.get("pages")
            items = res.get("list") or []
            for item in items:
                yield item
                got += 1
                if max_items is not None and got >= max_items:
                    return {"total": total, "pages": pages, "collected": got, "page": page}
            if not items:
                return {"total": total, "pages": pages, "collected": got, "page": page}
            if pages is not None and page >= int(pages):
                return {"total": total, "pages": pages, "collected": got, "page": page}
            page += 1

    @staticmethod
    def parse_link_ids(url: str) -> dict[str, str]:
        """从轮播/专区 link_url 解析 nav / goods。"""
        out: dict[str, str] = {}
        if not url:
            return out
        parsed = urlparse(url)
        qs = parse_qs(parsed.query)
        # hash query: #/selection?nav_type=1&id=51
        if not qs and "#" in url:
            frag = url.split("#", 1)[1]
            if "?" in frag:
                qs = parse_qs(frag.split("?", 1)[1])
        for key in ("id", "nav_id", "goods_id", "nav_type"):
            if key in qs and qs[key]:
                out[key] = qs[key][0]
        if "goods/detail" in url or "goods_id" in out:
            if "goods_id" not in out and "goods_id=" in url:
                out["goods_id"] = qs.get("goods_id", [""])[0]
        return out
