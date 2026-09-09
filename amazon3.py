import re
from openpyxl import load_workbook
from datetime import datetime
from zoneinfo import ZoneInfo
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from concurrent.futures import ThreadPoolExecutor, as_completed
import time
from queue import Queue
import tempfile
import psutil
import atexit
import threading
from loguru import logger
from lxml import html
import random
import pandas as pd

# # ============================ 配置区 =================================
# ASIN_dict  = {
#             '电即热':
#                 ['B002635ODW','B003XRC2IA','B01MR7Z39V','B01N5R9FQB','B01NAUZJPE','B0047V0KSU','B01MS9DVEE','B09393M76S','B0939BHX98','B0939T9X72','B09394DVP5','B0938YL6F5','B01N12VLI7','B01N35V1HH','B0CHRMSXRD','B0CHRNG2H4','B00529DDUI','B0DPZWVMHF','B0CVTQXW5Y','B0D2V1P149','B0D49Q2G8W','B0D49QRFGB','B0D2TY9N4M','B07GZ1XZXV','B07GYZ9MM4','B07GZ1WK29','B07GYZS4RF','B07GYZNS4D','B07GYYXXF8','B07GZ4PGXZ','B07GYZS4RJ','B07GYZ9MM2','B07GYQHWYJ','B07GZ33DKP','B07GZ3MFBX','B0D62WTN1M','B0C711Q56V','B0DGGH987T','B0DLKQY4B3','B0DLSL2FW3','B0BBT4ZR9R','B08MJ3NRG6','B08MJJHZ7J','B0F7WYM7QK','B0D8TY9T6S','B09394BZR7','B0FKB36VMJ','B0FKB5B8SG','B0FKB2D863','B0DGG6S814','B003FFXNOM','B01HTZD9M4','B000RYWZ48','B003FFVW6I','B00IX0V57Q','B0FL6MYLD2','B0FL73GY1F','B08ML63GSX','B0G9NBFR4D','B0FPQ677X7','B0FPQ7CS86','B0DWZM6W2W','B00IX1BRI2','B09CYFH5H7','B0C9LMNL4C','B0F9NPTK5F','B0FW4NSTFQ','B0GZGG3T4J','B0GT4W549G','B0GT4M317X'],
#              '燃热':
#                 ['B0FN7ZH4JF','B0FN84Q6GT','B0FN819PHN','B0FNWNTR55','B0FN819DC9','B0G633SC38','B0D8NW5HZ5','B0D8NSQQ3S','B0D8NL33M5','B0DJRM8ZL2','B0D8NQ4CCB','B0DHVM7V1F','B0DHVNMZHZ','B0D8NGTJNT','B094D42R1W','B0799JPXJV','B0DR158DJ2','B0DR17PMHH','B09D7D1BW2','B09D7L8NRM','B0DD438JP8','B0CNKBGLZ7','B0DD42P3B3','B0DD3W7XQV','B0CNK8YC52','B0G1YFTVZ2','B0GJFPY3DN','B0GJFPX2X3','B0GJFMHRTJ','B0GJFKZGDQ','B0CR1738ZD','B0C6PRDHF1','B0C6PYS9D2','B0GJFFMY5R','B0GJFJG5FL','B0FK9YJBL6','B0FKB9NVHY','B0GJG9B7N3','B0GSZ7FCRP','B0GSZFQQ2C','B0GSZKKKQJ','B0GSZT2NT2','B0FHCYGWQK','B0FHH8DGV7','B0FHH9RW9J','B0FHHB2RVD','B0CNQHT7CL','B07M65FJCQ','B07MH3C76M','B0B57NW13Y','B0CNQFDB5Z','B0B57ST42W','B0C8263551','B0C828QWTS','B0C8265TNX','B0C82611LN','B0C827FQPG','B0C827XPJ5','B0C828MJBK','B0C826LQ24','B0C826RQH9','B0C827G97F','B0C828BV45','B0C82F779S','B0C825Y8MN','B0C826LGV1','B0DRDBDS57','B0F9Y5ZNVF',],
#             '厨宝':
#                 ['B0148O658Y','B0148O65IE','B0148O64JE','B0FP5R9JSH','B0FP5P3W85','B0FP5R241K','B0B2Z5Z9FW','B0B2XBVTBG','B08BFXM5ML','B08BFL66GJ','B08BFKGYVP','B08BF78G8P','B0DKHGWLGB','B0CCDJKZHP','B0CCHPPWJC','B0CCJ3PS9J','B0FVRVX1R7','B0FJDCPYX4','B0CL9M23XX','B0D7MC2JCM','B0FN41NLLC','B0789DTY5L','B0789CC1TN','B07CV789V4','B0CQJL45KJ','B0D93DKPWJ','B08BJY39R4','B08BJY77WM','B08BJXMMNT','B08BJX98JB','B0DKHNPXHW','B083CVZ1DF','B0DBZL7ZT5','B0D6G8XTVK','B0DFCF5YGS','B0D6G4B16D','B0DBZQKFXX','B0D6GFL5JM','B0FDBD6V3C','B0FDB8WRCB','B0FDB7VS4N','B0FDB9PCQ8','B0FDB92Z99','B0FDB7CN35','B0FDG32Z9J','B0FDG3QY5H','B0FDG2WST7','B0FDG5FJ5S','B0FDG2XTQX','B0FDG3FG8N','B0FY6J48D8','B0FY6FWN51','B0FY6D722H','B0FY6F5WNQ','B0FF47TRBP','B0FF47M62P','B0FF47JGQT','B0GYP68DRG','B0GYP38C2V','B0GYPBMVG6','B0GYP58SMG','B0FF46Q9VX','B0GFK8S36H','B0GFJJG25X','B0GFJ62TW4','B0GFHLSBRK','B07RT71RYH','B07RS4VL8L','B07RT72MV6','B0D1K5WKJJ','B0D1K8CWMS','B0F47QBBCT','B0F47XL4D3','B0F483FV2P','B0GT73JPH3','B0GT7527K2','B0GWDDSVLM','B0F266PFZZ',]
#              }
import os
# ASIN_dict= {'错误':["B00IX1BRI2","B0FN84Q6GT","B0DJRM8ZL2","B0C6PYS9D2","B0C826RQH9","B0FP5R9JSH","B0DKHGWLGB","B0789CC1TN","B07CV789V4","B08BJXMMNT","B08BJX98JB"]}
ASIN_dict = {'1':['B0C6PVPN39','B0D25JSZN2','B0D21R1S39','B0C6DTBZ21','B0DB5S8Z55','B0D21R1S39','B0C6DTBZ21','B0DB5S8Z55',]}
ASIN_list   = [asin for asins in ASIN_dict.values() for asin in asins]
asin_to_cat = {asin: cat for cat, asins in ASIN_dict.items() for asin in asins}
zip_code = '90001'
file_path = r'D:\my\Excel\spider_man\亚马逊竞品信息.xlsx'
max_workers = 3
browser_pool = Queue(maxsize=max_workers)
zone_dict = {
    '10010': 'America/New_York',      # 美国东部（纽约）
    '90001': 'America/Los_Angeles',   # 美国西部（洛杉矶）
    '60601': 'America/Chicago',       # 美国中部（芝加哥）
}
# ===========================  第二配置 ==============================
_all_browsers = []
_browsers_lock = threading.Lock()
MAX_PAGES_PER_BROWSER = 15  #浏览器爬取每爬取15个ASIN就会杀掉重建
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
    # Edge 家族（同样带 window.chrome，指纹不冲突）
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/150.0.0.0 Safari/537.36 Edg/150.0.0.0",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/148.0.0.0 Safari/537.36 Edg/148.0.0.0",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/146.0.0.0 Safari/537.36 Edg/146.0.0.0",]

# ===================================================================


class TrackedBrowser:
    """包装浏览器，记录已处理页面数，到期自动回收"""
    def __init__(self):
        self.browser = create_browser()
        self.pages_used = 0

    def is_expired(self):
        return self.pages_used >= MAX_PAGES_PER_BROWSER

    def increment(self):
        self.pages_used += 1

def safe_quit(browser):
    """先优雅关闭，再强制清理"""
    try:
        browser.quit()
    except:
        pass
    force_kill(browser)

def recycle_browser(tracked):
    """杀掉旧浏览器，创建新的 TrackedBrowser 返回"""
    browser = tracked.browser
    with _browsers_lock:
        try:
            _all_browsers.remove(browser)
        except ValueError:
            pass
    safe_quit(browser)
    logger.info(f'♻ 浏览器已使用 {tracked.pages_used} 页，主动回收重建')
    return TrackedBrowser()

def handle_anti_bot_page(browser,url):
    """处理 Amazon 的反爬拦截页，点 Continue 后确认页面真正跳转"""
    max_attempts = 3

    for attempt in range(max_attempts):
        try:
            # 检测是否在拦截页上
            continue_btn = WebDriverWait(browser, 3).until(EC.element_to_be_clickable((By.XPATH,'//button[normalize-space()="Continue shopping"]')))
            if attempt > 0:
                logger.warning(f'检测到反爬拦截页，点击 Continue (第{attempt+1}次)')
            continue_btn.click()

            # 关键：等页面真正跳转离开拦截页
            WebDriverWait(browser, 10).until(
                EC.presence_of_element_located((By.ID, 'glow-ingress-line2'))
            )
            return True
        except:
            if 'Continue shopping' not in browser.page_source:
                return True  # 确认离开了拦截页
            time.sleep(1)

    return False  # 3次都没解决

def create_browser():
    """每个线程独立创建一个浏览器实例"""
    opt = Options()
    opt.add_argument('--headless=new')
    opt.add_argument('--no-sandbox')                                    # 禁用共享内存
    opt.add_argument('--disable-dev-shm-usage')                         # 禁用共享内存
    opt.add_argument('--blink-settings=imagesEnabled=false')            # 禁用图片
    opt.add_argument('--remote-debugging-pipe')                         # 避免端口冲突
    opt.add_argument('--disable-extensions')                            # 禁止加载所有浏览器扩展
    opt.add_argument('--disable-gpu')                                   # 禁用 GPU 硬件加速
    opt.add_argument('--disable-software-rasterizer')                   # 禁用软件光栅化
    opt.add_argument('--js-flags=--max-old-space-size=1024')             # 限制 V8 引擎的堆内存上限为 1024MB
    opt.add_argument(f'--user-data-dir={tempfile.mkdtemp(prefix="chrome_")}')
    opt.add_argument(f'user-agent={random.choice(USER_AGENTS)}')
    opt.add_experimental_option("excludeSwitches", ["enable-automation"])
    opt.add_experimental_option('useAutomationExtension', False)
    opt.page_load_strategy = 'eager'
    service = Service(executable_path=r'D:\my\chorme_of_some\chromedriver-win64\chromedriver.exe')
    browser = webdriver.Chrome(options=opt, service=service)
    browser.set_page_load_timeout(35)
    browser.execute_cdp_cmd('Page.addScriptToEvaluateOnNewDocument', {
        'source': '''
                Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
                Object.defineProperty(navigator, 'plugins', {get: () => [1, 2, 3, 4, 5]});
                Object.defineProperty(navigator, 'languages', {get: () => ['en-US', 'en']});
                window.chrome = { runtime: {} };
            '''
    })
    with _browsers_lock:  # ← 加锁
        _all_browsers.append(browser)
    return browser

def force_kill(browser):
    try:
        driver_proc = psutil.Process(browser.service.process.pid)
        for child in driver_proc.children(recursive=True):
            child.kill()
        driver_proc.kill()
    except (psutil.NoSuchProcess, psutil.AccessDenied):
        pass

def cleanup_all():
    with _browsers_lock:  # ← 加锁（遍历也需要保护）
        snapshot = list(_all_browsers)
    for b in snapshot:
        safe_quit(b)
    if snapshot:
        print(f'清理完毕 共 {len(snapshot)} 个浏览器')

# 第二步：注册到 atexit
atexit.register(cleanup_all)   # ← 这行必须在函数外面


def warm_up_pool():
    """预创建浏览器实例放入池"""
    for i in range(max_workers):
        browser_pool.put(TrackedBrowser())
    print('浏览器池预热完成')

def scrape_asin(asin, zip_code, retry=3):
    for i in range(retry):
        tracked = None
        try:
            tracked = browser_pool.get(timeout=60)
        except Exception:
            print(f'[超时] {asin} 第{i + 1}次获取浏览器超时')
            continue

        browser = tracked.browser
        result = None
        time.sleep(random.uniform(2, 5))
        try:
            result = get_info(browser, asin, zip_code)
            if result is None:
                raise ValueError(f'{asin} get_info 返回 None')
            tracked.increment()                        # 成功才计数
        except Exception as e:
            logger.warning(f'[重试] {asin} 第{i+1}次失败: {type(e).__name__}: {str(e)[:80]}')

        if result is not None:
            if tracked.is_expired():
                browser_pool.put(recycle_browser(tracked))
            else:
                try:
                    browser.get("about:blank")
                    browser_pool.put(tracked)
                except Exception:
                    # 清理失败，直接回收
                    browser_pool.put(recycle_browser(tracked))
            return asin, result
        else:
            try:
                browser.get("about:blank")
                tracked.increment()  # 也算一次使用
                if tracked.is_expired():
                    browser_pool.put(recycle_browser(tracked))
                else:
                    browser_pool.put(tracked)
            except Exception:
                # 浏览器级故障，才真正回收
                browser_pool.put(recycle_browser(tracked))

    return asin, None

def get_info(browser, ASIN, zip_code):
    # --------------打开商品页面-----------
    picture = None
    url = f'https://www.amazon.com/dp/{ASIN}'
    us_time = datetime.now(ZoneInfo(zone_dict[zip_code]))
    browser.get(url)
    time.sleep(1)
    if not handle_anti_bot_page(browser, url):
        raise RuntimeError('反爬拦截页无法绕过')
    wait = WebDriverWait(browser, 8)
    # -----------------------设置邮编-----------------------
    current_zipcode = wait.until(EC.presence_of_element_located((By.ID, 'glow-ingress-line2'))).text
    if zip_code not in current_zipcode:
        address_entry = wait.until(EC.element_to_be_clickable((By.ID, 'glow-ingress-block')))  # 收件地址
        address_entry.click()
        try:
            change_link = wait.until(EC.element_to_be_clickable((By.ID, 'GLUXChangePostalCodeLink')))  # change
            change_link.click()
        except:
            pass
        zip_input = wait.until(EC.element_to_be_clickable((By.ID, 'GLUXZipUpdateInput')))  # 元素在DOM+可见+可点击
        zip_input.clear()
        zip_input.send_keys(zip_code)
        apply_btn = wait.until(EC.element_to_be_clickable((By.ID, 'GLUXZipUpdate')))  # 点击aaply
        apply_btn.click()
        time.sleep(1)
        browser.get(url)
    # 等标题加载完，后获取中心html元素：centercol和right_col
    wait.until(EC.presence_of_element_located((By.XPATH, '//span[@id="productTitle"]')))
    center_col = wait.until(EC.presence_of_element_located((By.ID, 'centerCol')))
    right_col = wait.until(EC.presence_of_element_located((By.ID, 'rightCol')))
    center_html_str = center_col.get_attribute('innerHTML')
    center_tree = html.fromstring(center_html_str)  # 解析为 lxml 元素树
    right_html_str = right_col.get_attribute('innerHTML')
    right_tree = html.fromstring(right_html_str)  # 解析为 lxml 元素树
    # 使用xpath提取价格文本结算价/划线价/购物车价格
    settle_whole = center_tree.xpath('.//span[@id="apex-pricetopay-accessibility-label"]/..//span[@class="a-price-whole"]/text()')
    settle_frac = center_tree.xpath('.//span[@id="apex-pricetopay-accessibility-label"]/..//span[@class="a-price-fraction"]/text()')
    whole = settle_whole[0].strip() if settle_whole else ''
    frac = settle_frac[0].strip() if settle_frac else ''
    settle_price = f'{whole}.{frac}' if whole else None

    list_price_el = center_tree.xpath('.//span[contains(text(),"List Price: $") or contains(text(),"Typical price: $")]')
    if list_price_el:
        match = re.search(r'[\d,.]+', list_price_el[0].text_content())
        list_price = match.group() if match else None
    else:
        list_price = None

    regular_whole = right_tree.xpath('.//span[@class="a-text-bold" and contains(text(),"Regular Price")]/ancestor::div[@id="newAccordionCaption_feature_div"]/following-sibling::div[@id="apex_offerDisplay_desktop"]//span[@class="a-price-whole"]/text()')
    regular_frac = right_tree.xpath('.//span[@class="a-text-bold" and contains(text(),"Regular Price")]/ancestor::div[@id="newAccordionCaption_feature_div"]/following-sibling::div[@id="apex_offerDisplay_desktop"]//span[@class="a-price-fraction"]/text()')
    if regular_whole and regular_frac:
        regular_price = regular_whole[0].strip() + '.' + regular_frac[0].strip()
    else:
        regular_price = None
    # 标题 / 品牌 / 容量
    product_title = center_tree.xpath('.//span[@id="productTitle"]')[0].text.strip()
    brand_els = center_tree.xpath('.//tr[@class="a-spacing-small po-brand"]//span[@class="a-size-base po-break-word"]')
    brand = brand_els[0].text.strip() if brand_els and brand_els[0].text else None
    capacity_els = center_tree.xpath('.//tr[@class="a-spacing-small po-capacity"]//span[@class="a-size-base po-break-word"]')
    capacity = capacity_els[0].text.strip() if capacity_els and capacity_els[0].text else None
    # 获取活动信息
    promotion = {}
    coupon_text_el  = center_tree.xpath(".//span[contains(@id,'couponTextpctch')]")
    if coupon_text_el:
        coupon = re.search(r'(\d+%|\$[\d.]+)', coupon_text_el[0].text_content())
        promotion['coupon'] = coupon[0].strip() if coupon else None

    Free_bonus = center_tree.xpath("(.//label[text()='Free bonus item'])[1]")
    if Free_bonus:
        promotion['买赠'] = ""

    deal_el = center_tree.xpath( "//span[@id='dealBadgeSupportingText']")
    Limited_time_deal = deal_el[0].text_content().strip() if deal_el else None
    if Limited_time_deal == 'Prime Day Deal':
        promotion['PDBD'] = ''
    elif Limited_time_deal and ('Limited time deal' in Limited_time_deal or 'end' in Limited_time_deal.lower()):
        promotion['BD'] = ''

    Lowest_price_in_30_days = center_tree.xpath("//div[@class=' delightPricingBadge']/span[contains(text(),'Lowest price in 30 days')]")
    if Lowest_price_in_30_days:
        promotion['30天最低价'] = ""

    checkout_el  = center_tree.xpath("//div[contains(text(), 'Save') and contains(text(), 'at checkout')]")
    save_at_checkout = checkout_el[0].text_content().strip() if checkout_el else None
    if save_at_checkout:
        match = re.search(r'(\d+%)', save_at_checkout)
        promotion['code'] = match.group(1) if match else save_at_checkout.strip()

    Exclusive_Prime_price = center_tree.xpath("//span[@id='primeExclusivePricingMessage']/span[contains(text(),'Exclusive')]")
    if Exclusive_Prime_price:
        promotion['PED'] = ''

    save_code_el  = center_tree.xpath("//span[@class='promoPriceBlockMessage']/div/span/label[contains(text(),'Save')]")
    save_code = save_code_el[0].text_content().strip() if save_code_el and save_code_el[0].text_content() else None
    if save_code:
        promotion['code2'] = re.sub(r'Save |\s', '', save_code)

    buy_used = right_tree.xpath('//div[@id="usedBuySection"]')
    if buy_used:
        promotion['二手货'] = ""

    tiered_promotion = center_tree.xpath('//label[starts-with(@id, "greenBadge") and contains(text(),"off") and contains(text(),"Up to")]')
    if tiered_promotion:
        promotion['层级促销'] = ""

    stock = center_tree.xpath('.//div[@id="availability"]/span[contains(text(),"left in stock")]')
    if stock and not buy_used:
        promotion['缺货'] = ""

    lose_cart = right_tree.xpath('//span[@id="fod-cx-message-with-learn-more"]/span[text()="High price"]')
    if lose_cart and not regular_price:
        promotion['丢购物车'] = ''

    Frequently_returned_item =  center_tree.xpath('//span[@class and contains(text(),"Frequently returned item")]')
    if Frequently_returned_item:
        promotion['高退货'] = ''

    promotion_str = ' + '.join([f'{v}{k}' for k, v in promotion.items()]) if promotion else ''

    review_num_el = center_tree.xpath('//span[@id="acrCustomerReviewText"]')
    review_num = re.sub(r',|\(|\)', '', review_num_el[0].text) if review_num_el else ''

    score_el  = center_tree.xpath('//span[@id="acrPopover"]//span[@aria-hidden and @class]')
    score = score_el[0].text.strip() if score_el and score_el[0].text else ''

    # 大小类排名
    item_detais = wait.until(EC.element_to_be_clickable((By.XPATH, '//span[contains(text(),"Item details")]')))
    item_detais.click()
    big_kind_text = wait.until(EC.presence_of_element_located(
        (By.XPATH, '//th[contains(text(),"Best Sellers Rank")]/..//li[1]/span[@class="a-list-item"]'))).text
    match = re.search(r'#?([\d,]+)', big_kind_text)
    big_kind_rank = match.group(1).replace(',', '') if match else None
    small_kind_rank = wait.until(EC.presence_of_element_located(
        (By.XPATH, '//th[contains(text(),"Best Sellers Rank")]/..//li[2]/span[@class="a-list-item"]'))).text
    match = re.search(r'#?([\d,]+)', small_kind_rank)
    small_kind_rank = match.group(1).replace(',', '') if match else None

    # 要写入的所有信息
    product_info = [ASIN, picture,brand,capacity,settle_price, list_price, regular_price, promotion_str, big_kind_rank, small_kind_rank, score,
                    review_num,product_title,us_time.strftime('%Y-%m-%d'),us_time.strftime('%Y-%m-%d %H:%M:%S')]
    return product_info


if __name__ == '__main__':
    logger.info("神级脚本启动")
    wb = load_workbook(file_path)
    ws = wb['amazon222']
    warm_up_pool()
    results = {}
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_map = {executor.submit(scrape_asin, asin, zip_code): asin for asin in ASIN_list}
        done_count = 0
        for future in as_completed(future_map):
            asin, product_info = future.result()
            results[asin] = product_info
            done_count += 1
            status = '成功' if product_info else '失败'
            print(f'  [{done_count}/{len(ASIN_list)}] {asin} -> {status}')

    # -------- 按原始顺序写入 Excel（主线程操作，线程安全） --------
    failed_asinlist, rows = [], []
    for asin in ASIN_list:
        product_info = results.get(asin)
        if product_info:
            ws.append(product_info)
            rows.append(product_info)
        else:
            failed_asinlist.append(asin)
    print(f'成功抓取：{len(rows)}/{len(ASIN_list)} ')
    if failed_asinlist:
        print(f'失败ASIN列表：{','.join([f'"{asin}"' for asin in failed_asinlist])}')
    wb.save(file_path)

    today = pd.DataFrame(rows, columns=['ASIN', 'picture', 'brand', 'capacity', 'settle_price', 'list_price',
               'regular_price', 'promotion_str', 'big_kind_rank', 'small_kind_rank',
               'score', 'review_num', 'product_title', 'date', 'datetime'])
    today["ASIN"] = today["ASIN"].astype(str).str.strip().str.upper()
    today["类别"] = today["ASIN"].map(asin_to_cat).fillna("未匹配")

    now = datetime.now(ZoneInfo(zone_dict[zip_code])).strftime('%Y%m%d')
    out_dir = os.path.dirname(file_path)
    snap = os.path.join(out_dir, f'亚马逊竞品信息{now}.xlsx')
    today.to_excel(snap, index=False)
    logger.info(f'报告丰总运行结束，已导出桌面文件：亚马逊竞品信息{now}.xlsx')