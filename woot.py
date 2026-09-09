r"""周报合并脚本：一次运行完成 何老师(woot-services) + 巧豚豚(qtt.frual) 两份数据抓取，
写入同一个 Excel：D:\my\Excel\临时工作\周报\woot.xlsx
- sheet「何老师」：replace 覆盖
- sheet「巧豚豚」：overlay 追加
原文件 woot何老师2.py / 巧豚豚2.py 保持不变。
用法：
    python woot.py            # 依次跑两个爬虫
    python woot.py refresh    # 手动刷新巧豚豚 token（打开浏览器登录）
"""
import requests
import pandas as pd
import json
import os
import sys
import time
import random
from datetime import datetime, timedelta
from pathlib import Path
from urllib.parse import quote
from loguru import logger
from dotenv import load_dotenv

EXCEL_PATH = r'D:\my\Excel\临时工作\周报\woot.xlsx'

# ==================== 何老师（woot-services.com） ====================

load_dotenv()  # 从 .env 读取 WOOT_ACCOUNT / WOOT_PASSWORD
if not os.getenv("WOOT_PASSWORD"):
    logger.warning("⚠ WOOT_PASSWORD 为空！请检查 .env 文件是否存在、变量名是否正确，否则登录必失败")

WOOT_LOGIN_URL = "https://woot-services.com/api/auth:signIn"
WOOT_TOKEN_FILE = Path(__file__).parent / "token.json"


def woot_refresh_token() -> str:
    """调登录接口换新 JWT，写入 token.json，返回 token（不含 Bearer 前缀）。"""
    resp = requests.post(
        WOOT_LOGIN_URL,
        json={"account": os.getenv("WOOT_ACCOUNT", "TINA"),
              "password": os.getenv("WOOT_PASSWORD", "")},
        timeout=30,
    )
    if resp.status_code != 200:
        logger.error(f"登录失败 HTTP {resp.status_code}: {resp.text}")
        resp.raise_for_status()
    data = resp.json()
    token = (data.get("data") or {}).get("token")
    if not token:
        logger.error(f"登录响应中没有 token 字段: {data}")
        raise RuntimeError("登录接口返回异常，请检查 WOOT_ACCOUNT / WOOT_PASSWORD 是否正确")
    WOOT_TOKEN_FILE.write_text(
        json.dumps({"authorization": f"Bearer {token}"}, ensure_ascii=False),
        encoding="utf-8",
    )
    return token


def woot_get_authorization() -> str:
    """优先读本地 token.json；缺失或损坏则刷新。"""
    if WOOT_TOKEN_FILE.exists():
        try:
            return json.loads(WOOT_TOKEN_FILE.read_text(encoding="utf-8"))["authorization"]
        except Exception:
            pass
    return f"Bearer {woot_refresh_token()}"


def woot_authed_get(session, url, **kw):
    """带自动续期的 GET：返回 401/403（token 失效）时刷新并重试一次。"""
    headers = dict(kw.pop("headers", {}))
    headers["authorization"] = woot_get_authorization()
    r = session.get(url, headers=headers, **kw)
    if r.status_code in (401, 403):
        headers["authorization"] = f"Bearer {woot_refresh_token()}"
        r = session.get(url, headers=headers, **kw)
    return r


def run_he_laoshi():
    """何老师：全量爬取销量大于0的月度销售数据，replace 写入 sheet「何老师」。"""
    base_url = "https://woot-services.com/api/asin_sunmit:list"
    params = {
        "sort[]": "-entering_warehouse_time",
        "pageSize": 100,
        "filter": '{"$and":[{"createdBy":{"id":{"$eq":"{{$user.id}}"}}},{"$or":[{"asin_status":{"$eq":"db_submit"}},{"asin_status":{"$eq":"exit_warehouse_stage"}},{"asin_status":{"$eq":"wait_pay"}},{"asin_status":{"$eq":"finish"}}]}]}'
    }
    headers = {
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/149.0.0.0 Safari/537.36',
        'referer': 'https://woot-services.com/admin/illtlyh0tb8/popups/2cnop7e16d1/filterbytk/9409'
    }
    session = requests.Session()
    session.headers.update(headers)
    response = woot_authed_get(session, base_url, params=params, timeout=(10, 30))

    if response.status_code != 200:
        logger.error(f"[何老师] 主请求失败，状态码: {response.status_code},错误信息: {response.text}")
        return
    data = response.json()
    logger.info(f"[何老师] 请求成功! 返回数据条数: {len(data.get('data', []))}")
    extracted = [{"id": item["id"], "child_asin": item["child_asin"]} for item in data["data"]]

    filter_condition = {"sales_num": {"$gt": 0}}
    encoded_filter = quote(json.dumps(filter_condition))
    all_sales_data = []
    for idx, item in enumerate(extracted):
        asin_id = item['id']
        child_asin = item['child_asin']
        url = f'https://woot-services.com/api/asin_sunmit/{asin_id}/real_sales_id:list?sort[]=-sales_month&pageSize=20&filter={encoded_filter}'
        try:
            response = woot_authed_get(session, url, timeout=30)
            if response.status_code == 200:
                sales_data = response.json().get("data", [])
                for sale in sales_data:
                    all_sales_data.append({
                        "sales_month": sale.get("sales_month"),
                        "child_asin": child_asin,
                        "sales_num": sale.get("sales_num"),
                    })
                print(f"【{child_asin}】 [{idx+1}/{len(extracted)}]获取到 {len(sales_data)} 条销售记录")
            else:
                print(f"ASIN {child_asin} 请求失败，状态码: {response.status_code}")
        except Exception as e:
            logger.info(f"ASIN {child_asin} 请求异常: {str(e)[:50]}")
        time.sleep(0.6)
    logger.info(f'[何老师] 共计{len(all_sales_data)}条数据')
    df = pd.DataFrame(all_sales_data)
    df['数据来源'] = '何老师'
    df['sales_month'] = pd.to_datetime(df['sales_month'], errors='coerce').dt.tz_convert('Asia/Shanghai').dt.tz_localize(None).dt.date
    df = df.drop_duplicates(subset=['sales_month', 'child_asin', 'sales_num'])
    try:
        with pd.ExcelWriter(EXCEL_PATH, engine='openpyxl', if_sheet_exists='replace', mode='a') as writer:
            df.to_excel(writer, sheet_name='何老师', index=False)
    except PermissionError:
        logger.error("写入失败：woot.xlsx 正被 Excel 打开，请关闭后重跑")


# ==================== 巧豚豚（qtt.frual.com） ====================

QTT_TOKEN_CACHE = Path(__file__).parent / "qtt_token.txt"
QTT_USER_NAME = "admin@bc12821"
QTT_PASSWORD = "Fogatti123"
USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/150.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/149.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/148.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/146.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/145.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/144.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/143.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/142.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/150.0.0.0 Safari/537.36 Edg/150.0.0.0",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/148.0.0.0 Safari/537.36 Edg/148.0.0.0",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/146.0.0.0 Safari/537.36 Edg/146.0.0.0",
]


def qtt_fetch_authorization():
    """打开浏览器让用户登录(手动过验证码)，自动从网络请求抓取 authorization 并缓存。"""
    from playwright.sync_api import sync_playwright
    captured = {}
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()
        def on_request(req):
            auth = req.headers.get("authorization")
            if auth:
                captured["tok"] = auth
        page.on("request", on_request)
        page.goto("https://qtt.frual.com/app/#/order/sales-report", wait_until="domcontentloaded")
        try:
            page.wait_for_selector('xpath=//input[@placeholder="请输入账号"]', timeout=8000)
            page.fill('xpath=//input[@placeholder="请输入账号"]', QTT_USER_NAME)
            page.fill('xpath=//input[@placeholder="请输入密码"]', QTT_PASSWORD)
            page.click('xpath=//button[contains(@class,"login-btn")]', timeout=3000)
        except Exception as e:
            print(f"自动填表失败：{str(e)[:80]}，请手动登录")
        deadline = time.time() + 300
        while "tok" not in captured and time.time() < deadline:
            page.wait_for_timeout(500)
        browser.close()
    tok = captured.get("tok")
    if not tok:
        raise RuntimeError("未抓到 authorization，请确认是否成功登录。")
    QTT_TOKEN_CACHE.write_text(tok, encoding="utf-8")
    return tok


def qtt_normalize_token(tok):
    """规整 token：去空白，确保带 Bearer 前缀。"""
    tok = tok.strip()
    if not tok:
        return tok
    if not tok.lower().startswith("bearer "):
        tok = "Bearer " + tok
    return tok


def qtt_get_authorization(force_refresh=False):
    if not force_refresh and QTT_TOKEN_CACHE.exists():
        cached = QTT_TOKEN_CACHE.read_text(encoding="utf-8").strip()
        if cached:
            return qtt_normalize_token(cached)
    return qtt_fetch_authorization()


def qtt_get_sales(page_num=1):
    authorization = qtt_get_authorization()
    headers = {
        "User-Agent": f'{random.choice(USER_AGENTS)}',
        "authorization": authorization,
    }
    date_from = (datetime.now() - timedelta(days=2)).strftime('%Y-%m-%d')
    date_to = (datetime.now() - timedelta(days=2)).strftime('%Y-%m-%d')
    # date_from = '2026-01-01'
    # date_to = '2026-08-31'
    params = {
        "filter[reportDate][gte]": date_from,
        "filter[reportDate][lt]": date_to,
        "page": page_num,
        "per-page": 10
    }
    url = "https://qtt.frual.com/qtcapi/web/index.php/order/sales"

    def _parse(resp):
        data = resp.json()
        if not data.get('items'):
            print(f"状态码：{resp.status_code}，数据为空")
            return None, 1
        max_page = data['page']['pageCount']
        df = pd.DataFrame(data['items'])
        df['数据来源'] = '巧豚豚'
        df = df.loc[:, ['reportDate', '_childAsin', 'unitsSoldDailyAmazon', '数据来源']]
        return df, max_page

    response = requests.get(url, headers=headers, params=params)
    if response.status_code == 200:
        return _parse(response)
    elif response.status_code == 401:
        print("authorization 失效，尝试重新获取…")
        headers["authorization"] = qtt_fetch_authorization()
        response = requests.get(url, headers=headers, params=params)
        if response.status_code == 200:
            return _parse(response)
        return None, 1
    else:
        print(f"请求失败状态码：{response.status_code}")
        return None, 1


def run_qtt():
    """巧豚豚：抓最近一天销售报表，overlay 追加写入 sheet「巧豚豚」。"""
    all_data = []
    df_one, max_page = qtt_get_sales()
    all_data.append(df_one)
    if max_page > 1:
        for page_num in range(2, max_page + 1):
            df, _ = qtt_get_sales(page_num)
            all_data.append(df)
            time.sleep(0.5)
    all_data = [d for d in all_data if d is not None]
    if all_data:
        result = pd.concat(all_data, ignore_index=True)
        try:
            with pd.ExcelWriter(EXCEL_PATH, engine='openpyxl', if_sheet_exists='overlay', mode='a') as writer:
                ws = writer.book["巧豚豚"]
                result.to_excel(writer, sheet_name='巧豚豚', startrow=ws.max_row, index=False, header=False)
            print(f'[巧豚豚] 已写入到 {EXCEL_PATH} ')
        except PermissionError:
            logger.error("写入失败：woot.xlsx 正被 Excel 打开，请关闭后重跑")


# ==================== 主入口 ====================

if __name__ == '__main__':
    if len(sys.argv) > 1 and sys.argv[1] == "refresh":
        # 手动刷新巧豚豚 token：python woot.py refresh
        print("新 token:", qtt_fetch_authorization())
    else:
        run_he_laoshi()
        run_qtt()
        print("两个数据源全部跑完。")

        woot_path = r'D:\my\Excel\临时工作\周报\woot.xlsx'
        df_qtt = pd.read_excel(woot_path, sheet_name='巧豚豚')
        df_hetea = pd.read_excel(woot_path, sheet_name='何老师')
        df_hetea.rename(columns={'sales_month': '日期', 'child_asin': 'ASIN', 'sales_num': '销量'}, inplace=True)
        df_mjjl = pd.read_excel(woot_path, sheet_name='卖家精灵')
        df_mjjl.rename(columns={'日销量': '销量'}, inplace=True)
        df_mjjl['数据来源'] = '卖家精灵'
        df = pd.concat([df_qtt, df_hetea, df_mjjl])
        df['销量'] = pd.to_numeric(df['销量'], errors='coerce')
        df = df.loc[df['销量'] > 0, ['日期', 'ASIN', '销量', '数据来源']]
        df['日期'] = pd.to_datetime(df['日期'], errors='coerce').dt.date

        zs_file = r'D:\my\Excel\临时工作\中山份额\中山份额.xlsx'
        cs_file = r'D:\my\Excel\H10自动化\长沙份额.xlsx'
        df_zhongshan = pd.read_excel(zs_file,sheet_name='映射')
        df_zhongshan = df_zhongshan.rename(columns={'大类':'分类目'})
        df_changsha = pd.read_excel(cs_file,sheet_name='映射')
        mapping_df = pd.concat([df_zhongshan[['ASIN', '分类目', ]],df_changsha[['ASIN', '分类目']]]).drop_duplicates(subset=['ASIN'], keep='first')

        df = df.merge(mapping_df, on='ASIN', how='left')
        with pd.ExcelWriter(woot_path, engine='openpyxl', if_sheet_exists='replace', mode='a') as writer:
            df.to_excel(writer, sheet_name='合并结果', index=False)
        print(f'已写入到 {woot_path} ')
