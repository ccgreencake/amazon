def description(file_path,size_chart,part1,part2,part3):
    output_string = size_chart.replace("Size: ", "").replace(" Waist: ", " &emsp; ").replace(" Hip: ",
                                                                                             " &emsp; ").replace(
        " Length: ", " &emsp; ").replace("\n", " <br>\n")
    # 打开文本文件进行读取
    with open(file_path, 'r', encoding='utf-8') as file:
        # 读取文件内容
        file_content = file.read()
    # 进行替换操作
    file_content = file_content.replace("<strong>Length</strong><br>", "<strong>Length</strong><br>\n" + output_string)
    first_position = file_content.find("<p><strong></strong></p>")
    # 查找第二个 <p><strong> 的位置
    second_position = file_content.find("<p><strong></strong></p>", first_position + 1)
    # 查找第3个 <p><strong> 的位置
    third_position = file_content.find("<p><strong></strong></p>", second_position + 1)
    # 查找第一个 <p></p> 的位置
    first_p = file_content.find("<p></p>")
    # 查找第二个 <p></p>  的位置
    second_p = file_content.find("<p></p>", first_p + 1)
    # 查找第二个 <p></p>  的位置
    third_p = file_content.find("<p></p>", second_p + 1)

    product_description = file_content[:third_p] + "<p>" + part3[1] + "</p>" + file_content[third_p + 7:]
    product_description = product_description[:third_position] + "<p><strong>" + part3[0] + "</strong></p>" + product_description[third_position + 24:]
    product_description = product_description[:second_p] + "<p>" + part2[1] + "</p>" + product_description[second_p + 7:]
    product_description = product_description[:second_position] + "<p><strong>" + part2[0] + "</strong></p>" + product_description[second_position + 24:]
    product_description = product_description[:first_p] + "<p>" + part1[1] + "</p>" + product_description[first_p + 7:]
    product_description = product_description[:first_position] + "<p><strong>" + part1[0] + "</strong></p>" + product_description[first_position + 24:]
    return product_description

def rise(string):
    if "High" in string:
        rise = 'high'
    elif "Low" in string:
        rise = 'low'
    else:
        rise = 'mid'
    return rise

def item_type(string,feed_product):
    if feed_product == 'pants':
        if "Dress" in string:
            item_type = 'dress-pants'
        elif "Yoga" in string:
            item_type = 'yoga-pants'
        elif "Casual" in string:
            item_type = 'casual-pants'
        else:
            item_type = 'pants'
    elif feed_product == 'shorts':
        if "denim" in string or "jeans" in string:
            item_type = 'denim-shorts'
        elif "cargo" in string:
            item_type = 'cargo-shorts'
        elif "running" in string or "jogger" in string:
            item_type = 'running-shorts'
        elif "hiking" in string:
            item_type = 'hiking-shorts'
        else:
            item_type = 'shorts'
    elif feed_product == 'onepieceoutfit':
        item_type = 'jumpsuits-apparel'
    elif feed_product == 'overalls':
        item_type = 'overalls'
    return item_type

def feed_product(string):
    if 'Shorts' in string:
        feed_product = 'shorts'
    elif 'Overalls' in string:
        feed_product = 'overalls'
    elif 'Jumpsuit' in string:
        feed_product = 'onepieceoutfit'
    else:
        feed_product = 'pants'
    return feed_product

def leg_style(string):
    if "Flare" in string:
        leg_style = 'Flared'
    elif "Cropped" in string:
        leg_style = 'Cropped'
    else:
        leg_style = 'Wide'
    return leg_style