import requests

# 发送HTTP请求获取网页内容
url = "https://www.amz123.com/usatopkeywords/1?k=wide%20leg"  # 替换为你要爬取的目标网站的URL
response = requests.get(url)

# 检查请求是否成功
if response.status_code == 200:
    # 提取网页内容
    content = response.text

    # 打印网页内容
    print(content)
else:
    print("请求失败")