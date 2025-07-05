import requests
import pandas as pd

# 配置请求参数
api_url = "https://www.cwl.gov.cn/cwl_admin/front/cwlkj/search/kjxx/findDrawNotice"
params = {
    "name": "ssq",          # 双色球
    "issueCount": 30000,     # 足够大的数值覆盖目标期号范围
    "pageNo": 1,
    "pageSize": 30000,
    "systemType": "PC",
    "issueStart": "2003001",
    "issueEnd": "2025075"
}

# 发送请求（需添加浏览器头避免被拦截）
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
    "Referer": "https://www.cwl.gov.cn/ygkj/wqkjgg/"
}
response = requests.get(api_url, params=params, headers=headers)
data = response.json()

# 提取所需数据
results = []
for item in data["result"]:
    period = item["code"]          # 期号
    numbers = item["red"] + "," + item["blue"]  # 红球+蓝球（用逗号连接）
    results.append([period, numbers])

# 筛选2025013-2025075期的数据（需确认接口返回的期号范围）
filtered_data = [row for row in results if 2025013 <= int(row[0]) <= 2025075]

# 保存为CSV
df = pd.DataFrame(filtered_data, columns=["期号", "开奖号码"])
df.to_csv("ssq_results_2025013_2025075.csv", index=False, encoding='utf-8-sig')

print(f"成功保存 {len(filtered_data)} 期开奖数据到 ssq_results_2025013_2025075.csv")