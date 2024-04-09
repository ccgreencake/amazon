from bs4 import BeautifulSoup

# 假设html是你从网页中抓取到的HTML代码
html = '''
<div class="main-item lg:pl-[30px] pl-[10px]" style="background-color:;"
     data-v-a73e401d>
    <div class="col-start-1 col-end-3 overflow-hidden break-words"
         data-v-a73e401d><a title="wide leg pants woman"
                            class="text-[#1fab89] inline-block w-full break-words"
                            href="undefinedwide leg pants woman" target="_blank"
                            rel="nofollow" data-v-a73e401d><span
            data-v-a73e401d><span style='color:#fd7237'>wide leg</span> pants woman</span></a>
        <!----></div>
    <div class="grid grid-cols-3 text-[#6f6f6f] text-center col-start-3 col-end-8 relative"
         data-v-a73e401d><span data-v-a73e401d>1889</span><span data-v-a73e401d>2164</span>
        <div class="flex items-center justify-center" data-v-a73e401d><img
                src="https://static.amz123.com/pb_assets/up.c81a040b.svg" alt=""
                data-v-a73e401d><span class="ml-[5px]"
                                      data-v-a73e401d>275</span></div>
    </div>
    <a href="undefinedwide leg pants woman" target="_blank" rel="nofollow"
       class="col-start-8 col-end-9 mx-auto" data-v-a73e401d><img
            src="https://img.amz123.com/static/images/hot_words/google.svg"
            class="cursor-pointer" data-v-a73e401d></a>
</div>
<div class="main-item lg:pl-[30px] pl-[10px]" style="background-color:#f9f9f9;"
     data-v-a73e401d>
    <div class="col-start-1 col-end-3 overflow-hidden break-words"
         data-v-a73e401d><a title="wide leg jeans woman"
                            class="text-[#1fab89] inline-block w-full break-words"
                            href="undefinedwide leg jeans woman" target="_blank"
                            rel="nofollow" data-v-a73e401d><span
            data-v-a73e401d><span style='color:#fd7237'>wide leg</span> jeans woman</span></a>
        <!----></div>
    <div class="grid grid-cols-3 text-[#6f6f6f] text-center col-start-3 col-end-8 relative"
         data-v-a73e401d><span data-v-a73e401d>7353</span><span data-v-a73e401d>7219</span>
        <div class="flex items-center justify-center" data-v-a73e401d><img
                src="https://img.amz123.com/static/images/hot_words/down.svg"
                alt="" data-v-a73e401d><span class="ml-[5px]" data-v-a73e401d>134</span>
        </div>
    </div>
    <a href="undefinedwide leg jeans woman" target="_blank" rel="nofollow"
       class="col-start-8 col-end-9 mx-auto" data-v-a73e401d><img
            src="https://img.amz123.com/static/images/hot_words/google.svg"
            class="cursor-pointer" data-v-a73e401d></a>
</div>
'''

# 使用BeautifulSoup解析HTML
soup = BeautifulSoup(html, 'html.parser')

# 根据位置找到目标元素
items = soup.find_all('div', class_='main-item')

# 提取指定位置的文本
results = []
for item in items:
    text_span = item.find('span', style='color:#fd7237')
    if text_span:
        result = text_span.get_text(strip=True)
        results.append(result)

# 打印结果
for result in results:
    print(result)