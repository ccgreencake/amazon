from 新上架 import DataClean
from 新上架.Combination import TitleCombiner
from 新上架.DataFill import AmazonTemplateFiller




# type = "SWEATSHIRT"
# type = "OVERALLS"
# type = "DRESS"
# type = "ONE_PIECE_OUTFIT"
# type = "SHIRT"
# type = "SKIRT"
# type = "SHORTS"
type = "PANTS"
brand='YOULE'
char_prefix='T'
if char_prefix == 'T':
    price = 42.49
    shipping = 0
    list_price = price
else:
    price = 9.99
    shipping = 4.99
    list_price = price + shipping
# 自定义sku
sku_gen = DataClean.SKUGenerator(base_dir=r'F:\product')
raw_df = sku_gen.run_sku_process(
    file_index=1,
    char_prefix=char_prefix,
    brand=brand
)
total_rows = len(raw_df)  # 列数
# 打印看看sku对不对
# print(raw_df['custom_sku'])
# 性别
gender = "men"
# gender = "women"

if gender == "men":
    department_name = "Mens"
    target_gender = "Male"
else:
    department_name = "Womens"
    target_gender = "Female"

style = "Casual"
life_style = "Casual"
age_range_description = "Adult"
special_size = "Big & Tall"
pocket_description = "Utility Pocket"
Brand = "Bakgeerle"
cibiao = "men tactical pants"
attributes = [
                "baggy",
                "casual",
                "plus size",
                "comfy",
                "soft",
                "comfortable",
                "elegant",
                "cozy",
                "trendy",
                "loose fit",
                "summer",
                "lightweight"
                "oversized",
                "breathable",
                "petite",
                "y2k",
                "big and tall"
                ]

title_list = raw_df['标题'].tolist()
colors_series = raw_df['颜色']
item_name = TitleCombiner().create_combined_titles(
                num=total_rows,
                cibiao=cibiao,
                title_list=title_list,
                colors_series=colors_series,
                flag="NORMAL"
            )
st_results = TitleCombiner().create_search_terms(
                num=total_rows,
                cibiao=cibiao,  # 确保你 Excel 有这个 Sheet
                title_list=title_list,
                attr_pool=attributes,
                flag="NORMAL"
        )
template_path = "description_template.txt"  # 需要自己建一个这文件测试

# 模拟创建一下测试 HTML 模板
with open(template_path, "w", encoding="utf-8") as f:
    f.write("<p><strong>⚠️Important Notice:</strong></p>\n")
    f.write("<p>Before making a purchase...</p>\n")
    f.write("<p><strong>Size Chart:</strong></p>\n<p>\n</p>\n")
    f.write("<p><strong>Note:</strong> recommend choosing larger size.</p>")
# 1. 提取倒数第三列 (-3)
# ==========================================
# iloc[:, -3] 的意思是：取所有行(:)，取倒数第三列(-3)
bullet_series_2 = raw_df.iloc[:, -3]

# ==========================================
# 2. 清洗数据，提取出干净的 5 个点
# ==========================================
# dropna() 去掉所有的 nan，然后转化为纯 list
clean_bullets = bullet_series_2.dropna().tolist()

# 截取前 5 个（防止原表有多余的杂乱数据）
p5_list = clean_bullets[:5]
product_description = TitleCombiner().create_html_description(
            file_path=template_path,
            size_chart=raw_df["尺码表"][0],
            p5=p5_list
        )
# 1. 定义你想要的列名排列顺序 (2, 3, 1, 4, 5, 6, 7, 8, 9)
# ==========================================
target_image_cols = [
    '代理链接 2', '代理链接 3', '代理链接 1',
    '代理链接 4', '代理链接 5', '代理链接 6',
    '代理链接 7', '代理链接 8', '代理链接 9'
]

# ==========================================
# 2. 安全过滤：防止原始表格缺列报错
# （比如有的供应商表格只有 7 张图，没有 8 和 9）
# ==========================================
# 这一步会按 target_image_cols 的顺序，挑出原表中真实存在的列
valid_cols = [col for col in target_image_cols if col in raw_df.columns]

if not valid_cols:
    print("警告：在表格中没有找到任何'代理链接'相关的列！")
else:
    # ==========================================
    # 3. 提取并重排数据
    # ==========================================
    # 只要传入列名列表，Pandas 就会自动按照这个列表的顺序生成一个新的二维 DataFrame
    reordered_images_df = raw_df[valid_cols]

    # 如果你需要标准的 Python 二维列表结构 (List of Lists)，可以加上这一句：
    # 结果示例: [ ['url2', 'url3', 'url1', ...], ['url2', 'url3', 'url1', ...] ]
    images_2d_list = reordered_images_df.values.tolist()
swatch_image_url = raw_df["代理链接 2"]

# 各种写死的列

listing_action = ["updateCreate or Replace (Full Update)"] * total_rows
product_type = [type] * total_rows
brand_name = [Brand] * total_rows
quantity = "1000"



# 模板路径
template_path = r'C:\Users\Administrator\Desktop\new_yo_template.xlsx'

# 实例化 (只写一次路径)
filler = AmazonTemplateFiller(original_path=template_path)

# 调用填充方法 (只传列名和列表)
# 它会自动去第 4 行找这些字眼，然后从第 7 行往下填
# 填写 自定义sku
filler.fill_column_data("SKU", raw_df['custom_sku'])
# 填写类目
filler.fill_column_data("Product Type", product_type)
filler.fill_column_data("Generic Keywords",st_results)

# 填充颜色
filler.fill_column_skip_first("Color Map",colors_series)
filler.fill_column_skip_first("Color",raw_df['自定义颜色'])

# 填写价格和运费
filler.fill_2d_data_skip_first('List Price',[str(list_price)] * total_rows)
filler.fill_2d_data_skip_first('Your Price USD (Sell on Amazon, US)', [str(price)] * total_rows)
filler.fill_2d_data_skip_first('Merchant Shipping Group (US)', [str(shipping)] * total_rows)

# 填充大描述
filler.fill_column_data("Product Description",[product_description] * total_rows)



size_system = ["US"] * total_rows
size_class = ["Alpha"] * total_rows
body_type = ["Regular"] * total_rows
height_type = ["Regular"] * total_rows
size = raw_df["亚马逊尺寸"]
# 1. 定义类型到分类的映射关系
type_to_category = {
    "SWEATSHIRT": "Apparel",
    "DRESS":      "Apparel",
    "OVERALLS":   "Bottoms",
    "SHORTS":     "Bottoms",
    "PANTS":      "Bottoms",
    "PANT":       "Bottoms",  # 兼容你之前的写法
    "SHIRT":      "Shirt",
    "SKIRT":      "Skirt",
}

# 2. 获取当前 type 对应的分类名
# 使用 .get() 可以防止 type 不在字典里时程序报错
category = type_to_category.get(type.upper())

# 3. 如果找到了对应的分类，则执行填充逻辑
if category:
    # 亚马逊的列名规律非常统一：{分类名} + {属性名}
    filler.fill_column_skip_first(f"{category} Size System", size_system)
    filler.fill_column_skip_first(f"{category} Size Class", size_class)
    filler.fill_column_skip_first(f"{category} Body Type", body_type)
    filler.fill_column_skip_first(f"{category} Height Type", height_type)
    filler.fill_column_skip_first(f"{category} Size Value", size)
else:
    print(f"⚠️ 警告：未定义的类型 '{type}'，请检查映射表。")
# 特殊情况
if type == "ONE_PIECE_OUTFIT":
    filler.fill_column_skip_first("Size", size)


parentage_level = ["child"] * total_rows
parent_sku = raw_df['custom_sku'][0]
parentage_level[0] = "parent"
item_condition = "New"
variation_theme_name = ["SIZE/COLOR"] * total_rows
import_designation = "Imported"
product_id_type = ["GTIN Exempt"] * total_rows
fulfillment_channel_code = "DEFAULT"
item_package_length = "25"
package_length_unit = "Centimeters"
item_package_width = "20"
package_width_unit = "Centimeters"
item_package_height = "2"
package_height_unit = "Centimeters"
package_weight = "150"
package_weight_unit = "Grams"
country_of_origin = "China"
item_weight = "150"
item_weight_unit = "Grams"
care_instructions = "Machine Wash"
pattern = "Solid"
unit_count = "1"
unit_count_type = "Count"
closure_type = "Elastic Closure"
rise_style = "Mid Rise"
leg_style = "Wide"
weave_type = "Knit"
filler.fill_column_data("Listing Action", listing_action)
filler.fill_column_data("Parentage Level", parentage_level)
filler.fill_column_skip_first("Parent SKU", [parent_sku] * total_rows)
filler.fill_column_data("Variation Theme Name", variation_theme_name)
filler.fill_column_data("Item Name", item_name)
filler.fill_column_data("Brand Name", brand_name)
filler.fill_column_data("Product Id Type", product_id_type)
filler.fill_column_data("Model Number",brand_name)
filler.fill_column_data("Model Name",brand_name)
filler.fill_column_data("Manufacturer",brand_name)
filler.fill_2d_data_skip_first("Main Image URL",images_2d_list)
filler.fill_column_skip_first("Swatch Image URL",swatch_image_url)
filler.fill_column_skip_first("Style" ,[style] * total_rows)
filler.fill_column_skip_first("Lifestyle" ,[life_style] * total_rows)
filler.fill_column_skip_first("Department Name" ,[department_name] * total_rows)
filler.fill_column_data("Target Gender" ,[target_gender] * total_rows)
filler.fill_column_data("Age Range Description" ,[age_range_description] * total_rows)


# 连续快速填写
# 定义单行的模板数据
material_row = ["Polyester", "Polyester", "Polyester", "71%Polyester,18%Cotton,11%Spandex","1","1"]
# 使用列表推导式，直接生成具有 total_rows 行的二维矩阵
material_2d_list = [material_row for _ in range(total_rows)]
filler.fill_2d_data_backward("Item Package Quantity",material_2d_list)
filler.fill_column_skip_first("Special Size",[special_size] * total_rows)
filler.fill_column_skip_first("Pocket Description",[pocket_description] * total_rows)
filler.fill_column_data("Item Condition", [item_condition] * total_rows)
filler.fill_column_data("Import Designation", [import_designation] * total_rows)
filler.fill_column_data("Quantity (US)", [quantity] * total_rows)
filler.fill_column_data("Fulfillment Channel Code (US)", [fulfillment_channel_code] * total_rows)
filler.fill_column_data("Item Package Length", [item_package_length] * total_rows)
filler.fill_column_data("Package Length Unit", [package_length_unit] * total_rows)
filler.fill_column_data("Item Package Width", [item_package_width] * total_rows)
filler.fill_column_data("Package Width Unit", [package_width_unit] * total_rows)
filler.fill_column_data("Item Package Height", [item_package_height] * total_rows)
filler.fill_column_data("Package Height Unit", [package_height_unit] * total_rows)
filler.fill_column_data("Package Weight", [package_weight] * total_rows)
filler.fill_column_data("Package Weight Unit", [package_weight_unit] * total_rows)
filler.fill_column_data("Country of Origin", [country_of_origin] * total_rows)
filler.fill_column_data("Item Weight", [item_weight] * total_rows)
filler.fill_column_data("Item Weight Unit", [item_weight_unit] * total_rows)
filler.fill_column_data("Care Instructions", [care_instructions] * total_rows)
filler.fill_column_data("Pattern", [pattern] * total_rows)
filler.fill_column_data("Unit Count", [unit_count] * total_rows)
filler.fill_column_data("Unit Count Type", [unit_count_type] * total_rows)
filler.fill_column_data("Closure Type" , [closure_type] * total_rows)
filler.fill_column_data("Rise Style" , [rise_style] * total_rows)
filler.fill_column_data("Leg Style" , [leg_style] * total_rows)
filler.fill_column_data("Weave Type" , [weave_type] * total_rows)






# 再加一道保险：如果原表连 5 个点都不够，用空字符串补齐 5 个位置，防止错位
while len(p5_list) < 5:
    p5_list.append("")

print(f"✅ 成功提取五点模板：\n1. {p5_list[0]}\n2. {p5_list[1]}...")

# ==========================================
# 3. 广播数据：将这 5 个点复制给每一行，形成二维列表
# ==========================================
total_rows = len(raw_df)

# 使用列表推导式，生成一个拥有 total_rows 行的二维列表
# 每一行长得一模一样，都是这 5 个五点描述
bullets_2d_list = [p5_list for _ in range(total_rows)]

# ==========================================
# 4. 填充入亚马逊表格
# ==========================================
filler.fill_2d_data_backward("Bullet Point", bullets_2d_list)

# 最后统一保存
filler.save_and_close()