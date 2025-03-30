import http.client
import requests
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
import http
import json
import urllib.parse
import time

# 獲取第一個公司ID資訊表
company_info_url1 = "https://openapi.twse.com.tw/v1/opendata/t187ap03_L"
response1 = requests.get(company_info_url1)
company_info_data1 = response1.json()

# 提取第一個公司ID資訊表的公司ID和公司名稱
company_id_name_map1 = {entry["公司代號"]: entry["公司名稱"]
                        for entry in company_info_data1}

# 獲取第二個公司ID資訊表
company_info_url2 = "https://www.tpex.org.tw/openapi/v1/mopsfin_t187ap03_O"
response2 = requests.get(company_info_url2)
company_info_data2 = response2.json()

# 提取第二個公司ID資訊表的公司ID和公司名稱
company_id_name_map2 = {entry["SecuritiesCompanyCode"]
    : entry["CompanyName"] for entry in company_info_data2}

# 找出兩個公司ID來源的重複公司
duplicate_companies = set(company_id_name_map1.keys()
                          ) & set(company_id_name_map2.keys())

# 顯示重複公司的名稱與編號
print("重複公司列表:")
for company_id in duplicate_companies:
    company_name1 = company_id_name_map1.get(company_id, "N/A")
    company_name2 = company_id_name_map2.get(company_id, "N/A")
    print(
        f"公司ID: {company_id}, 公司名稱(第一個來源): {company_name1}, 公司名稱(第二個來源): {company_name2}")

# 合併兩個公司ID和公司名稱的映射
company_id_name_map = {**company_id_name_map1, **company_id_name_map2}

total_companies = len(company_id_name_map)
print(f"總計 {total_companies} 家公司")

# 初始化一個空的列表來存儲所有公司的資訊
all_companies_data = []

# 獲取當前日期
current_date = datetime.now()

# 計算5年前的日期
five_years_later = current_date - timedelta(days=365*5)  # 假設一年為365天，不考慮閏年情況

# 格式化開始日期（5年前的日期）
start_date_string = five_years_later.strftime('%Y%m%d')

# 格式化結束日期（當前日期）
end_date_string = current_date.strftime('%Y%m%d')


# 打印開始日期和結束日期
print("開始日期:", start_date_string)
print("結束日期:", end_date_string)

# 初始化參數: 最小資料量 (日)
min_cdates = 60

# 迭代所有公司ID
for idx, (company_id, company_name) in enumerate(company_id_name_map.items(), start=1):
    # 顯示提示訊息
    print(
        f"處理第 {idx}/{total_companies} 家公司 - 公司代號: {company_id}, 公司名稱: {company_name}")

    # 使用統一的URL來查找公司資訊
    url = "https://mops.twse.com.tw/mops/web/ezsearch_query"

    # 設定 POST 請求所需的參數
    payload = {
        'step': '00',
        'RADIO_CM': '2',
        'TYPEK': '',
        'CO_MARKET': '',
        'CO_ID': company_id,
        'PRO_ITEM': 'F22',
        'SUBJECT': '',
        'SDATE': start_date_string,
        'EDATE': '',
        'lang': 'TW',
        'AN': '',
    }

    # 將 payload 轉換為 x-www-form-urlencoded 格式
    encoded_payload = urllib.parse.urlencode(payload)

    data = None
    retry_later = True

    while retry_later:
        # 發送 POST 請求獲取資料
        retry_later = False
        response = requests.post(url, data=encoded_payload, headers={
            'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8'})

        # 解析 JSON 格式的回應
        try:
            # 去除 UTF-8 BOM，然後解析響應內容
            text_response = response.text.strip('\ufeff')
            data = json.loads(text_response)
        except http.client.RemoteDisconnected as http_ex:
            print('網路發生錯誤，等待 5 秒後重試...')
            retry_later = True
            time.sleep(5)
        except Exception as e:
            if 'Overrun' in text_response:
                print('短時間發送太多請求，等待 5 秒後重試...')
                retry_later = True
                time.sleep(5)

    # 處理資料
    try:
        # 提取所需資訊，以字典形式存儲
        if data['status'] == 'success':
            cdates = [entry["CDATE"] for entry in data['data']]

            cdates_len = len(cdates)
            if cdates_len < 48:
                print(f"數據量: {cdates_len}; 小於 48, 跳過此公司")
                continue

            min_cdates = min(cdates_len, min_cdates)

            company_info = {
                '公司名稱': company_name,
                '公司代號': company_id,
                '資訊': cdates
            }
            all_companies_data.append(company_info)
        else:
            print(f"查詢失敗，跳過此公司 ({company_id}); {', '.join(data['message'])}")
    except Exception as e:
        print(f"發生錯誤，跳過此公司 ({company_id}), error: {e}")

# 只保留最小資料量
print(f"發布最少數量為：{min_cdates}, 將所有多於最少數量的資料清除...")
for comp_data in all_companies_data:
    cdates = comp_data['資訊']
    comp_data['資訊'] = cdates[len(cdates) - min_cdates:]

# 將數據輸出為JSON文件
output_file = 'all_companies_f22.json'
with open(output_file, 'w', encoding='utf-8') as json_file:
    json.dump(all_companies_data, json_file, ensure_ascii=False, indent=2)

print(f"所有公司資訊已成功提取並保存到 {output_file}")
