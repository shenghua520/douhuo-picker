"""抖货商城（微唯宝）选品接口客户端。

会话保持：Playwright 持久化浏览器上下文（user_data_dir），登录一次后长期复用，
不需要手工导出/导入 Cookie。
接口调用：全部走页面内 fetch，天然携带 Cookie/Referer/Origin，与前端行为一致。

接口与字段来自对 /choice/v3/js/app.*.js 及懒加载分包的逆向，非猜测。
仅实现抓数必需的能力，不做通用 SDK。
"""

from __future__ import annotations

import json
import random
import time
from pathlib import Path
from typing import Any

from playwright.sync_api import Page, sync_playwright

BASE = "https://www.douhuomall.com"
ENTRY_URL = BASE + "/choice/v3/#/index"        # 选品首页（用户给的入口）
LOGIN_URL = BASE + "/choice/v3/#/register_login"

API_GOODS_LIST = "/manager/Choice/getGoodsList"
API_TOP_CATE = "/manager/Choice/getTopCate"
API_BRAND = "/manager/Choice/getBrand"
API_SUPPLY = "/manager/Choice/getSupply"
API_USER_INFO = "/manager/Choice/getUserInfo"

# getUserInfo 里 sys_id == 11 表示「试用账号」：能看，但商品/价格不全
TRIAL_SYS_ID = "11"
STATE_OK, STATE_TRIAL, STATE_NONE = "已登录", "试用账号", "未登录"

# 前端 axios 响应拦截器里与登录态相关的错误码
ERR_NOT_LOGIN = {1014, 1020}
ERR_LOGIN_REDIRECT = {1011}

# 选品页排序：1 利润率 / 2 零售价 / 3 成本价；sort_desc：1 降序 / 2 升序
SORT_FIELD = {"profit": 1, "retail": 2, "cost": 3}
SORT_DESC = {"desc": 1, "asc": 2}
SORT_KEYS = {1: "profit", 2: "retail_price", 3: "plat_price"}  # 本地重排用的字段

# 商品详情链接模板
GOODS_URL = BASE + "/choice/v3/#/goods?id={goods_id}&nav_type={nav_type}"


class ApiError(RuntimeError):
    """接口返回非 0 错误码。"""

    def __init__(self, code: int, msg: str, url: str):
        super().__init__(f"{url} -> error_code={code} {msg}")
        self.code, self.msg, self.url = code, msg, url


class NotLoggedIn(RuntimeError):
    """登录态失效。"""


class PageLoadError(RuntimeError):
    """页面加载/接口不可达。"""


# 页面内 fetch：返回原始文本交给 Python 解析，避免大对象跨界序列化开销
_FETCH_JS = """
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


class DouhuoClient:
    """抖货商城选品接口客户端。用法：with DouhuoClient(...) as c: c.search(...)"""

    def __init__(
        self,
        profile_dir: Path,
        headless: bool = True,
        timeout: int = 30,
        retries: int = 3,
        min_delay: float = 0.6,
        max_delay: float = 1.4,
    ):
        self.profile_dir = Path(profile_dir)
        self.headless = headless
        self.timeout_ms = timeout * 1000
        self.retries = retries
        self.min_delay, self.max_delay = min_delay, max_delay
        self._pw = self._ctx = self._page = None
        self._dict_cache: dict[str, list[dict]] = {}

    # ---------- 生命周期 ----------

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
        self.open_entry()
        return self

    def __exit__(self, *exc) -> None:
        for closer in (self._ctx, self._pw):
            if closer:
                try:
                    closer.close() if closer is self._ctx else closer.stop()
                except Exception:
                    pass

    # ---------- 页面 ----------

    def open_entry(self, url: str = ENTRY_URL) -> None:
        """打开选品页并等待前端挂载。加载失败按指数退避重试。"""
        last = None
        for attempt in range(1, self.retries + 1):
            try:
                self._page.goto(url, wait_until="domcontentloaded", timeout=self.timeout_ms)
                self._page.wait_for_function(
                    "() => document.readyState === 'complete'", timeout=self.timeout_ms
                )
                self._page.wait_for_timeout(1500)  # 等 Vue 挂载 + 首屏接口
                return
            except Exception as exc:  # noqa: BLE001 - 网络/渲染抖动统一重试
                last = exc
                time.sleep(attempt * 2)
        raise PageLoadError(f"选品页加载失败（重试 {self.retries} 次）：{last}")

    # ---------- 接口 ----------

    def _post(self, path: str, payload: dict) -> dict:
        """POST 一个 JSON 接口，返回 result 字段。"""
        last: Exception | None = None
        for attempt in range(1, self.retries + 1):
            self._throttle()
            try:
                raw = self._page.evaluate(
                    _FETCH_JS, [path, payload, self.timeout_ms]
                )
            except Exception as exc:  # noqa: BLE001
                last = PageLoadError(f"{path} 请求异常：{exc}")
                time.sleep(attempt * 2)
                continue
            if not raw.get("ok"):
                last = PageLoadError(f"{path} HTTP {raw.get('status')} {raw.get('text', '')[:200]}")
                time.sleep(attempt * 2)
                continue
            try:
                data = json.loads(raw["text"])
            except json.JSONDecodeError:
                # 被重定向到登录页 / 网关返回 HTML
                raise NotLoggedIn(f"{path} 返回非 JSON（多半是登录态失效）") from None
            code, msg = data.get("error_code"), data.get("error_msg", "")
            if code in ERR_NOT_LOGIN or code in ERR_LOGIN_REDIRECT:
                raise NotLoggedIn(msg or "登录态失效")
            if code:
                raise ApiError(code, msg, path)
            return data.get("result") or {}
        raise PageLoadError(f"{path} 连续 {self.retries} 次失败：{last}")

    def _throttle(self) -> None:
        time.sleep(random.uniform(self.min_delay, self.max_delay))

    def login_state(self) -> str:
        """已登录 / 试用账号 / 未登录。"""
        try:
            info = self._post(API_USER_INFO, {})
        except NotLoggedIn:
            return STATE_NONE
        except Exception:  # noqa: BLE001 - 网络抖动不算未登录，交给上层重试
            return STATE_OK
        return STATE_TRIAL if str(info.get("sys_id")) == TRIAL_SYS_ID else STATE_OK

    def ensure_login(self, wait_minutes: int = 10) -> str:
        """真正未登录时引导人工登录（需有头模式）；试用账号只提示、不阻断。"""
        state = self.login_state()
        if state != STATE_NONE:
            return state
        if self.headless:
            raise NotLoggedIn(
                "登录态失效且当前为无头模式。请先执行：python run.py --login "
                "（会打开浏览器，扫码/短信登录后关闭窗口即可）"
            )
        print("\n>>> 未登录。请在弹出的浏览器中完成登录（短信验证码/扫码）。")
        print(f">>> 登录页：{LOGIN_URL}\n")
        try:
            self._page.goto(LOGIN_URL, wait_until="domcontentloaded", timeout=self.timeout_ms)
        except Exception:
            pass
        deadline = time.time() + wait_minutes * 60
        while time.time() < deadline:
            time.sleep(3)
            if self.login_state() != STATE_NONE:
                print(">>> 登录成功，继续采集。")
                self.open_entry()
                return self.login_state()
        raise NotLoggedIn(f"{wait_minutes} 分钟内未检测到登录成功")

    # ---------- 业务 ----------

    def search(
        self,
        keyword: str = "",
        *,
        nav_type: int = 0,
        nav_id: str = "",
        cate_id: int = 0,
        brand_id: list[int] | None = None,
        supply_type: list[int] | None = None,
        min_plat: float = 0,
        max_plat: float = 0,
        min_profit: float = 0,
        max_profit: float = 0,
        is_freight: bool = False,
        sort_field: int = 0,
        sort_desc: int = 0,
        page_size: int = 20,
        max_items: int = 60,
        max_pages: int = 20,
    ) -> list[dict]:
        """按条件检索商品，自动翻页，最多返回 max_items 条。"""
        payload = {
            "nav_type": nav_type,
            "nav_id": nav_id,
            "type": 1 if keyword else 0,      # 1=按商品名 2=按货号
            "spu_sn": "",
            "spu_name": keyword or "",
            "sort_field": sort_field,
            "sort_desc": sort_desc,
            "page": 1,
            "limit": page_size,
            "supply_type": supply_type or [],
            "cate_id": cate_id,
            "brand_id": brand_id or [],
            "tag_id": "",
            "min_profit": min_profit,
            "max_profit": max_profit,
            "min_plat": min_plat,
            "max_plat": max_plat,
            "is_freight": 1 if is_freight else "",
        }
        items: list[dict] = []
        for page in range(1, max_pages + 1):
            payload["page"] = page
            result = self._post(API_GOODS_LIST, payload)
            items.extend(result.get("list") or [])
            if len(items) >= max_items or page >= (result.get("pages") or 1):
                break

        # 已知坑：服务端对利润率（计算字段）排序不生效，零售价/成本价也只是近似有序，
        # 统一在本地重排一次，保证「按利润率降序取前 N」这类需求结果确定。
        key = SORT_KEYS.get(sort_field)
        if key and sort_desc in (1, 2):
            items.sort(key=lambda g: _to_float(g.get(key)), reverse=sort_desc == 1)
        return items[:max_items]

    # ---------- 名称 -> ID ----------

    def _dict(self, key: str, loader) -> list[dict]:
        if key not in self._dict_cache:
            self._dict_cache[key] = loader()
        return self._dict_cache[key]

    def categories(self) -> list[dict]:
        """一级 + 二级类目，摊平成 [{id, name}]。"""

        def load():
            out = []
            for top in self._post(API_TOP_CATE, {}) or []:
                out.append({"id": top.get("cate_id"), "name": top.get("cate_name")})
                for child in top.get("child") or []:
                    out.append({"id": child.get("cate_id"), "name": child.get("cate_name")})
            return out

        return self._dict("cate", load)

    def brands(self, cate_id: int = 0, max_pages: int = 8) -> list[dict]:
        """品牌字典（全量约 3700 个，按 cate_id 缓存，只在填了「品牌」列时才拉取）。"""
        key = f"brand:{cate_id}"
        if key not in self._dict_cache:
            out = []
            for page in range(1, max_pages + 1):
                res = self._post(API_BRAND, {
                    "page": page, "limit": 500, "group_id": 0,
                    "cate_id": cate_id, "navid": 0,
                })
                out += [{"id": b.get("brand_id"), "name": b.get("brand_name")}
                        for b in res.get("list") or []]
                if page >= (res.get("pages") or 1):
                    break
            self._dict_cache[key] = out
        return self._dict_cache[key]

    def supplies(self) -> list[dict]:
        return self._dict("supply", lambda: [
            {"id": s.get("id"), "name": s.get("name")} for s in self._post(API_SUPPLY, {}) or []
        ])

    @staticmethod
    def resolve_id(value: Any, entries: list[dict], label: str) -> int:
        """支持直接填 ID，或填名称（模糊匹配）。空值返回 0。"""
        text = str(value or "").strip()
        if not text:
            return 0
        if text.isdigit():
            return int(text)
        for entry in entries:
            if text == str(entry.get("name") or ""):
                return int(entry.get("id") or 0)
        for entry in entries:  # 先精确后包含，避免「纸」命中一堆
            if text in str(entry.get("name") or ""):
                return int(entry.get("id") or 0)
        raise ValueError(f"无法把「{text}」解析为{label}ID：请改填数字 ID，或检查名称是否正确")

    @classmethod
    def resolve_ids(cls, value: Any, entries: list[dict], label: str) -> list[int]:
        """同 resolve_id，但支持用 / 或 、分隔填多个（品牌、渠道是多选字段）。"""
        parts = [p for p in str(value or "").replace("、", "/").split("/") if p.strip()]
        ids = [i for i in (cls.resolve_id(p, entries, label) for p in parts) if i]
        return sorted(set(ids))


def normalize_goods(raw: dict, nav_type: int = 0) -> dict:
    """把接口原始商品对象摊平成写 Excel 的一行字典。"""
    plat = _to_float(raw.get("plat_price"))
    retail = _to_float(raw.get("retail_price"))
    reduce_lo, reduce_hi = _to_float(raw.get("min_reduce")), _to_float(raw.get("max_reduce"))
    return {
        "goods_id": str(raw.get("goods_id") or ""),
        "spu_name": str(raw.get("spu_name") or "").strip(),
        "supply_name": str(raw.get("supply_name") or "").strip(),
        "plat_price": plat,
        "retail_price": retail,
        "purchase_price": (
            f"{reduce_lo:g} - {reduce_hi:g}" if reduce_lo != reduce_hi else f"{reduce_hi:g}"
        ) if (reduce_lo or reduce_hi) else "",
        "profit": _to_float(raw.get("profit")),
        "commission": round(retail - plat, 2) if retail and plat else "",  # 零售价 - 代发价
        "is_freight": "包邮" if raw.get("is_freight") == 1 else "",
        "sort": max(0, int(_to_float(raw.get("sort")))),  # 负值按前端逻辑归零
        "status": "售罄" if str(raw.get("status")) == "0" else "在售",
        "url": GOODS_URL.format(goods_id=raw.get("goods_id"), nav_type=nav_type),
        "main_img": str(raw.get("main_img") or "").strip(),
    }


def _to_float(value: Any) -> float:
    try:
        return float(str(value).strip())
    except (TypeError, ValueError):
        return 0.0
