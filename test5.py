def replace_at_position(text, position, replacement):
    return text[:position] + replacement + text[position + len(replacement):]


# 打开文本文件进行读取
with open('html.txt', 'r', encoding='utf-8') as file:
    # 读取文件内容
    file_content = file.read()

# 查找第一个 <p><strong> 的位置
first_position = file_content.find("<p><strong></strong></p>")
if first_position != -1:
    print("第一个 <p><strong> 出现的位置：", first_position)
else:
    print("未找到第一个 <p><strong>")

# 查找第二个 <p><strong> 的位置
second_position = file_content.find("<p><strong></strong></p>", first_position + 1)
if second_position != -1:
    print("第二个 <p><strong> 出现的位置：", second_position)
else:
    print("未找到第二个 <p><strong>")

# 查找第3个 <p><strong> 的位置
third_position = file_content.find("<p><strong></strong></p>", second_position + 1)
if second_position != -1:
    print("第三个 <p><strong> 出现的位置：", third_position)
else:
    print("未找到第三个 <p><strong>")

# 查找第一个 <p></p> 的位置
first_p = file_content.find("<p></p>")
if first_p != -1:
    print("第一个 <p></p> 出现的位置：", first_p)
else:
    print("未找到第一个 <p></p>")

# 查找第二个 <p></p>  的位置
second_p = file_content.find("<p></p>", first_p + 1)
if second_p != -1:
    print("第二个 <p><strong> 出现的位置：", second_p)
else:
    print("未找到第二个 <p></p>")

# 查找第二个 <p></p>  的位置
third_p = file_content.find("<p></p>", second_p + 1)
if second_p != -1:
    print("第三个 <p></p> 出现的位置：", third_p)
else:
    print("未找到第三个 <p></p>")

new_text = file_content[:third_p] + "<p>556asdfasdasdasdasdasdascas</p>" + file_content[third_p+7:]
new_text = new_text[:third_position] + "<p><strong>556</strong></p>" + new_text[third_position+24:]
new_text = new_text[:second_p] + "<p>adsfadsfasdf</p>" + new_text[second_p+24:]
new_text = new_text[:second_position] + "<p><strong>123</strong></p>" + new_text[second_position+24:]
new_text = new_text[:first_p] + "<p>asdfasdfasdfsdsdsddddddd</p>" + new_text[first_p+7:]
new_text = new_text[:first_position] + "<p><strong>abc</strong></p>" + new_text[first_position+24:]





# 将替换后的内容写入新的文本文件
with open('output.txt', 'w', encoding='utf-8') as file:
    file.write(new_text)
