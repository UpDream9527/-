# 主要用于H10下载竞品数据，,目前每天需要更新令牌x-pacvue-token
import requests
import pandas as pd
import time
from datetime import datetime, timedelta
import random

def get_asin(headers,page=1, idea='174696'):
    asin_url = f'https://h10api.pacvue.com/rta/product-launchpad/v1/ideas/products?accountId=1545497626&ideaId={idea}&page={page}&perPage=100&sort=-title'
    response = requests.get(url=asin_url, headers=headers, timeout=(10, 30))
    if response.status_code == 401:
        raise Exception("Token已过期，请更新x-pacvue-token (状态码401)")
    if response.status_code != 200:
        raise Exception(f"请求失败，状态码: {response.status_code}")

    df = pd.DataFrame(response.json()['data'])
    ASIN_list = df['asin'].tolist()
    max_page = response.json()['meta']['totalPages']
    time.sleep(random.uniform(0.9, 2.2))
    return ASIN_list, max_page,idea
def retry_request(func, *args, max_retries=3, retry_delay=5, **kwargs):
    """
    通用重试包装器
    :param func: 要执行的函数
    :param max_retries: 最大重试次数（默认3次）
    :param retry_delay: 重试间隔秒数（默认5秒）
    """
    for attempt in range(1, max_retries + 1):
        try:
            result = func(*args, **kwargs)
            return result
        except Exception as e:
            print(f"  ⚠ 第 {attempt}/{max_retries} 次尝试失败: {e}")
            if attempt < max_retries:
                wait = retry_delay + random.uniform(1, 3)
                print(f"  ⏳ {wait:.1f} 秒后重试...")
                time.sleep(wait)
            else:
                raise  # 全部重试用完，向上抛出异常

def fetch_child_data_batch(asins, date_from, date_to, batch_size=10,max_round_retries=3):
    """
    分批获取childSales和childRevenue数据
    参数:
        asins: ASIN列表
        date_from: 开始日期，格式'YYYY-MM-DD'
        date_to: 结束日期，格式'YYYY-MM-DD'
        batch_size: 每批处理的ASIN数量，默认10
    """
    all_data = []
    total_asins = len(asins)
    failed_batches = []
    # 将ASIN列表分批
    for i in range(0, total_asins, batch_size):
        batch = asins[i:i + batch_size]
        batch_num = i // batch_size + 1
        try:
            df_batch = retry_request(fetch_single_batch, batch, date_from, date_to, headers,max_retries=3, retry_delay=5)
            all_data.append(df_batch)
            print(f"第 {batch_num} 批完成，获取 {len(df_batch)} 行数据")
            if i + batch_size < total_asins:
                time.sleep(random.uniform(1.1, 3.9))
        except Exception as e:
            print(f"第 {batch_num} 批失败: {e}")
            failed_batches.append((batch_num, batch))
    for round_num in range(1, max_round_retries):
        if not failed_batches:
            break
        print(f"\n🔄 第 {round_num + 1} 轮重试，剩余 {len(failed_batches)} 个失败批次...")
        time.sleep(random.uniform(8, 15))  # 轮次间等待更久
        still_failed = []
        for batch_num, batch in failed_batches:
            try:
                df_batch = retry_request(fetch_single_batch, batch, date_from, date_to, headers,
                                         max_retries=2, retry_delay=8)
                all_data.append(df_batch)
                print(f"  ✅ 第 {batch_num} 批重试成功，获取 {len(df_batch)} 行数据")
                time.sleep(random.uniform(2, 5))
            except Exception as e:
                print(f"  ❌ 第 {batch_num} 批重试仍失败: {e}")
                still_failed.append((batch_num, batch))
        failed_batches = still_failed

            # ---------- 汇总结果 ----------
    if failed_batches:
        print(f"\n⚠ 最终仍有 {len(failed_batches)} 批失败: {[b[0] for b in failed_batches]}")
    if all_data:
        return pd.concat(all_data, ignore_index=True)
    else:
        print("所有批次均失败，未获取到任何数据")
        return pd.DataFrame()


def fetch_single_batch(asins, date_from, date_to,headers):
    """
    单次请求函数
    """
    url = "https://h10api.pacvue.com/rta/product-launchpad/v1/products/spark-line"
    params = {
        "accountId": 1545497626,
        "asins[]": asins,
        "dateFrom": date_from,
        "dateTo": date_to,
        "marketplace": "ATVPDKIKX0DER"
    }
    response = requests.get(url, params=params, headers=headers, timeout=(10, 60))
    if response.status_code == 401:
        raise Exception(f"Token已过期，请更新x-pacvue-token (状态码401)")
    if response.status_code != 200:
        raise Exception(f"API请求失败，状态码: {response.status_code}，响应: {response.text[:200]}")
    data = response.json()['data']

    # 转换为DataFrame
    df = pd.DataFrame([
        {'日期': date,'ASIN': asin,
         'childSales': asin_data['childSales'].get(date),
         'childRevenue': asin_data.get('childRevenue', {}).get(date)}
        for asin, asin_data in data.items()
        for date in asin_data.get('childSales', {})
    ])
    time.sleep(random.uniform(1.1, 2.4))
    return df


if __name__ == '__main__':
    # USER_AGENTS = [
    #     "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/150.0.0.0 Safari/537.36",
    #     "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/149.0.0.0 Safari/537.36",
    #     "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/148.0.0.0 Safari/537.36",
    #     "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36",
    #     "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/146.0.0.0 Safari/537.36",
    #     "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/145.0.0.0 Safari/537.36",
    #     "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/144.0.0.0 Safari/537.36",
    #     "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/143.0.0.0 Safari/537.36",
    #     "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/142.0.0.0 Safari/537.36",
    #     # Edge 家族（同样带 window.chrome，指纹不冲突）
    #     "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/150.0.0.0 Safari/537.36 Edg/150.0.0.0",
    #     "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/148.0.0.0 Safari/537.36 Edg/148.0.0.0",
    #     "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/146.0.0.0 Safari/537.36 Edg/146.0.0.0", ]
    USER_AGENTS = [
        # ===== Chrome (Windows) =====
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0.0.0 Safari/537.36",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",

        # ===== Edge (Windows) =====
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36 Edg/130.0.0.0",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36 Edg/129.0.0.0",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36 Edg/128.0.0.0",

        # ===== Firefox (Windows) — 独立指纹，适合轮换 =====
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:132.0) Gecko/20100101 Firefox/132.0",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:131.0) Gecko/20100101 Firefox/131.0",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:130.0) Gecko/20100101 Firefox/130.0",

        # ===== Safari (macOS) — 用于伪装 Mac 用户 =====
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 14_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.5 Safari/605.1.15",
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 14_6) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.4 Safari/605.1.15",

        # ===== Safari (iOS) — 移动端指纹 =====
        "Mozilla/5.0 (iPhone; CPU iPhone OS 18_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/18.0 Mobile/15E148 Safari/604.1",
        "Mozilla/5.0 (iPhone; CPU iPhone OS 17_6 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.6 Mobile/15E148 Safari/604.1",
    ]


    type_dict = {
        '燃热': '188653', '净水': '178155', '厨宝': '174696', '电即热': '122909',
        '取暖器': '193044','房车空调': '193046', '房车厨电': '197935','房车热水器': '193747',
        # 全量
        # '燃热':'211125','净水':'211124','厨宝':'211122','电热':'211121'
        # ,'取暖':'211120','空调':'211119','房热':'211118',

    }
    headers = {
        'user-agent': f'{random.choice(USER_AGENTS)}',
        'x-pacvue-token': 'Bearer eyJhbGciOiJSUzI1NiIsImtpZCI6IjZGNEIxQ0Y5NENFNTE0M0M5MUQ2MjhCQTJGRjEyNEM1NkQ0QjdCRkIiLCJ0eXAiOiJKV1QiLCJ4NXQiOiJiMHNjLVV6bEZEeVIxaWk2TF9Fa3hXMUxlX3MifQ.eyJuYmYiOjE3ODg5MTYwMjIsImV4cCI6MTc4OTAwMjQyMiwiaXNzIjoiaHR0cDovL2lkZW50aXR5IiwiYXVkIjpbImh0dHA6Ly9pZGVudGl0eS9yZXNvdXJjZXMiLCJhcGkxIiwid2ljayJdLCJjbGllbnRfaWQiOiJjbGllbnQuaDEwLmp3dCIsInN1YiI6IjE1NDU5MzEwMDciLCJhdXRoX3RpbWUiOjE3ODg1MDM5MjEsImlkcCI6ImxvY2FsIiwidXNlcklkIjoiMTU0NTkzMTAwNyIsInJvbGUiOiJVc2VyIiwidXNlckluZm8iOiJ7XCJtYWlsXCI6XCJwb3d0ZWt6aG9uZ3NoYW5Ab3V0bG9vay5jb21cIixcInVzZXJOYW1lXCI6XCJBY2NvdW50NFwiLFwidXNlclJvbGVcIjpcIlVzZXJcIixcInVzZXJJZFwiOjUwNTYwMDQ5LFwicm9vdFVzZXJJZFwiOjUwMDA2NDM3LFwiY2xpZW50SWRcIjo1MDAwNjA4NyxcImlzU3VwZXJBZG1pblwiOmZhbHNlLFwiaDEwVXNlcklkXCI6MTU0NTkzMTAwNyxcImgxMEFjY291bnRJZFwiOjE1NDU0OTc2MjYsXCJvcmlnaW5DbGllbnRJZFwiOm51bGx9Iiwic2NvcGUiOlsib3BlbmlkIiwiYXBpMSIsIndpY2siLCJvZmZsaW5lX2FjY2VzcyJdLCJhbXIiOlsiYmFzaWNhdXRodG9rZW4iXX0.TTPxt4PwZp0xBX0OLMeZj8FUckgr8ap1buZkoVRIhvbpfS79A03-_Zv0nsSmJvgGPCbBUj8MA3KQgIP_Lf8Y2Is4ZcScckq6KEYu-YIytZRD0GUR_LVXfd6ZWAzJS76NqBQXiDRILq_tyod--SOspbNQzLB3XGm4NuK__vvILF9Etg1Wst0QAZPrmYDqBDEg6zzz2TTTg46w5Hg_baPwRNByGANKg0FclDTpbmJAluNaTZwMCAzUnRQq4QrLWPw6uz0uf8obLhJJXIOlmKl70CCpRSBCIn09aOEpb2ZL61WOFsPD9NeTV1Du_DnsARIj2gDDM8XNlSanSVv31ch4SQ'
    }
    asin_dict = {}
    asin_to_type = {}
    all_df = []
    for type ,value in type_dict.items():               # 循环各分类
        type_asins = []
        ASIN_list_one, max_page, idea = retry_request(get_asin, headers, page=1, idea=value, max_retries=4, retry_delay=8)
        type_asins.extend(ASIN_list_one)
        for page_num in range(2, max_page + 1):         # 分类页数
            asin_list, _, idea = retry_request(get_asin, headers, page=page_num, idea=idea, max_retries=4, retry_delay=8)  # 关键：传入idea_id
            type_asins.extend(asin_list)
        asin_dict[type] = type_asins
        for a in type_asins:  # ★新增：登记每个ASIN所属分类
            asin_to_type.setdefault(a, []).append(type)
        print(f'【{type}】获取完成，共{len(type_asins)}个ASIN')
    print(asin_dict)

    # ------------------以上为获取ASIN，下面为获取销量销售额------------------------
    date_to = (datetime.now() - timedelta(days=2)).strftime('%Y-%m-%d')
    date_from = (datetime.now() - timedelta(days=3)).strftime('%Y-%m-%d')
    # date_from = '2025-08-01'
    # date_to = '2025-08-31'
    # 对上面获取到的ASIN去重
    all_asin = set()
    for asin_list in asin_dict.values():
        all_asin.update(asin_list)
    all_asin = list(all_asin)
    print(f"去重后共有 {len(all_asin)} 个唯一ASIN,日期范围: {date_from} 至 {date_to}")
    # 分批获取数据
    df_final = fetch_child_data_batch(all_asin, date_from, date_to, batch_size=83)
    out_file = r'C:\Users\fgt\Desktop\child_data22.xlsx'
    if not df_final.empty:
        df_final['数据来源'] = '竞品'
        df_final['品类'] = df_final['ASIN'].map(lambda x: '、'.join(asin_to_type.get(x, ['未知'])))
        df_cs = df_final.loc[df_final['品类'].isin(['燃热','净水','电即热','厨宝'])]
        df_zs = df_final.loc[df_final['品类'].isin(['房车热水器','房车空调','房车厨电','取暖器'])]
        with pd.ExcelWriter(out_file, engine='openpyxl', if_sheet_exists='replace',mode='a') as writer:
            df_final.to_excel(writer, sheet_name='py2', index=False)
        with pd.ExcelWriter(out_file, engine='openpyxl', if_sheet_exists='replace',mode='a') as writer:
            df_cs.to_excel(writer, sheet_name='长沙', index=False)
        with pd.ExcelWriter(out_file, engine='openpyxl', if_sheet_exists='replace',mode='a') as writer:
            df_zs.to_excel(writer, sheet_name='中山', index=False)
        print(f"✅ 成功获取 {len(df_final)} 行数据，已保存到 {out_file}")
    else:
        print("❌ 未获取到任何数据，请检查Token和网络连接")