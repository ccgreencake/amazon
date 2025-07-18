import requests
from bs4 import BeautifulSoup

def fetch_amazon_page_with_cookies(url):
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
                      'AppleWebKit/537.36 (KHTML, like Gecko) Chrome/97.0.4692.71 Safari/537.36',
        'Accept-Language': 'en-US,en;q=0.9',
        'Accept-Encoding': 'gzip, deflate, br'
    }
    # 将你在浏览器获取的 Cookie 信息填入下面的字典（仅作示例，实际需合法获取并维护）
    cookies = {
        'session-id': 'YOUR_SESSION_ID',
        'session-token': 'YOUR_SESSION_TOKEN',
        # 其他必要的 cookie
    }
    response = requests.get(url, headers=headers, cookies=cookies, timeout=10)
    if response.status_code == 200:
        soup = BeautifulSoup(response.content, 'lxml')
        print(soup.prettify())
        return soup
    else:
        print(f"请求失败，状态码：{response.status_code}")
        return None

if __name__ == '__main__':
    url = "https://www.amazon.com/gp/product/B0DNMPYRBY?th=1&psc=1"
    fetch_amazon_page_with_cookies(url)
