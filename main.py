import requests
from bs4 import BeautifulSoup
import json
import re


def parse_curl_command(curl_command):
    """
    解析 cURL 命令，提取 cookies 和 user-agent（適配大小寫不敏感的 headers）
    """
    cookies = {}
    user_agent = ""

    # 正則提取所有 headers
    headers = re.findall(r"-H '([^:]+): (.*?)'", curl_command)

    # 將 headers 的鍵轉為小寫，以適配大小寫不敏感
    headers_dict = {key.lower(): value for key, value in headers}

    return headers_dict


def get_multiline_input(prompt):
    """
    支援多行輸入，根據行尾的 '\' 自動判斷輸入是否結束
    """
    print(prompt)
    lines = []
    while True:
        try:
            line = input()
            if line.endswith("\\"):
                lines.append(line[:-1])  # 移除行尾的 '\' 並繼續下一行
            else:
                lines.append(line)
                break  # 無 '\' 表示輸入結束
        except EOFError:
            break
    return "".join(lines)


# 提示使用者提供多行 cURL 命令
curl_command = get_multiline_input(
    "請先在 Chrome/Edge 手動開啟 https://www.iqvalue.com/Frontend/stock/shareholding?stockId=1102 並通過驗證，\n"
    "然後開啟 Chrome/Edge 的開發者工具 (F12) 並重新整理網頁，\n"
    "之後於 Network 分頁中右鍵點擊第一項請求，選擇 'Copy > Copy as cURL (bash)'，\n"
    "然後將複製的命令貼上這裡並按下 ENTER："
)

# 解析 cURL 命令
request_headers = parse_curl_command(curl_command)

# 測試解析結果
print("解析成功！以下是提取的資訊：")
print(f"Headers: {request_headers}")

# 獲取第一個公司ID資訊表
company_info_url1 = "https://openapi.twse.com.tw/v1/opendata/t187ap03_L"
response1 = requests.get(company_info_url1)
company_info_data1 = response1.json()

# 提取第一個公司ID資訊表的公司ID和公司名稱
company_id_name_map1 = {entry["公司代號"]: entry["公司名稱"] for entry in company_info_data1}

# 獲取第二個公司ID資訊表
company_info_url2 = "https://www.tpex.org.tw/openapi/v1/mopsfin_t187ap03_O"
response2 = requests.get(company_info_url2)
company_info_data2 = response2.json()

# 提取第二個公司ID資訊表的公司ID和公司名稱
company_id_name_map2 = {entry["SecuritiesCompanyCode"]: entry["CompanyName"] for entry in company_info_data2}

# 找出兩個公司ID來源的重複公司
duplicate_companies = set(company_id_name_map1.keys()) & set(company_id_name_map2.keys())

# 顯示重複公司的名稱與編號
print("重複公司列表:")
for company_id in duplicate_companies:
    company_name1 = company_id_name_map1.get(company_id, "N/A")
    company_name2 = company_id_name_map2.get(company_id, "N/A")
    print(f"公司ID: {company_id}, 公司名稱(第一個來源): {company_name1}, 公司名稱(第二個來源): {company_name2}")

# 合併兩個公司ID和公司名稱的映射
company_id_name_map = {**company_id_name_map1, **company_id_name_map2}

total_companies = len(company_id_name_map)
print(f"總計 {total_companies} 家公司")

# 初始化一個空的列表來存儲所有公司的資訊
all_companies_data = []

# 迭代所有公司ID
for idx, (company_id, company_name) in enumerate(company_id_name_map.items(), start=1):
    # 顯示提示訊息
    print(f"處理第 {idx}/{total_companies} 家公司 - 公司代號: {company_id}, 公司名稱: {company_name}")

    # 使用統一的URL來查找公司資訊
    url = f"https://www.iqvalue.com/Frontend/stock/shareholding?stockId={company_id}"

    # 發送 GET 請求，帶入 headers 和 cookies
    response = requests.get(url, headers=request_headers)
    response.encoding = 'utf-8'  # 手動設定編碼

    # 使用 Beautiful Soup 解析 HTML 內容
    soup = BeautifulSoup(response.text, 'html.parser')

    # 找到包含董監事持股信息的表格
    table = soup.find('table', class_='radius')

    # 初始化一個空的列表來存儲表格數據
    data = []
    try:
        # 提取表頭信息
        headers = [header.text.strip() for header in table.find_all('th')]

        # 提取每一行的數據
        for row in table.find_all('tr')[1:]:
            row_data = [cell.text.strip() for cell in row.find_all('td')]
            entry = dict(zip(headers, row_data))

            # 在每筆資料的開頭加入公司名稱和公司代號
            entry['公司名稱'] = company_name
            entry['公司代號'] = company_id

            data.append(entry)

        # 整合數據到所有公司的列表中
        all_companies_data.extend(data)
    except Exception as e:
        print(f"發生錯誤，跳過此公司 ({company_id}), error: {e}")

# 將數據輸出為 JSON 文件
output_file = 'all_companies_info.json'
with open(output_file, 'w', encoding='utf-8') as json_file:
    json.dump(all_companies_data, json_file, ensure_ascii=False, indent=2)

print(f"所有公司資訊已成功提取並保存到 {output_file}")
