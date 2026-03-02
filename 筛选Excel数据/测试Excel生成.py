import pandas as pd
import numpy as np
from datetime import datetime, timedelta

# 生成销售明细数据
np.random.seed(42)  # 固定随机种子，保证数据可重复
dates = [datetime(2023, 1, 1) + timedelta(days=i) for i in range(180)]  # 2023年1-6月
product_categories = ["电子产品", "家居用品", "服装鞋帽", "食品饮料"]
products = {
    "电子产品": ["智能手机", "笔记本电脑", "平板电脑", "耳机"],
    "家居用品": ["沙发", "床", "衣柜", "桌椅"],
    "服装鞋帽": ["T恤", "裤子", "鞋子", "帽子"],
    "食品饮料": ["零食", "饮料", "水果", "肉类"]
}
regions = ["华东", "华北", "华南", "西部"]
customer_types = ["个人", "企业"]

# 生成100行数据
data = []
for _ in range(100):
    date = np.random.choice(dates)
    category = np.random.choice(product_categories)
    product = np.random.choice(products[category])
    region = np.random.choice(regions)
    sales = np.random.randint(100, 10000)  # 销售额100-10000元
    quantity = np.random.randint(1, 20)  # 销量1-20件
    customer_type = np.random.choice(customer_types)
    data.append([
        date.strftime("%Y-%m-%d"), category, product, region, sales, quantity, customer_type
    ])

# 创建DataFrame
sales_df = pd.DataFrame(
    data,
    columns=["订单日期", "产品类别", "产品名称", "销售地区", "销售额（元）", "销量（件）", "客户类型"]
)

# 地区编码表
region_df = pd.DataFrame({
    "销售地区": ["华东", "华北", "华南", "西部"],
    "地区编码": ["EC001", "NC002", "SC003", "WC004"],
    "区域经理": ["张三", "李四", "王五", "赵六"]
})

# 产品信息表
product_info = []
for category in products:
    for product in products[category]:
        cost = np.random.randint(50, 5000)  # 成本价50-5000元
        supplier = f"{category[:2]}供应商{np.random.randint(1, 10)}"
        product_info.append([product, supplier, cost])

product_df = pd.DataFrame(
    product_info,
    columns=["产品名称", "供应商", "成本价（元）"]
)

# 保存到Excel
with pd.ExcelWriter("sales_data.xlsx", engine="openpyxl") as writer:
    sales_df.to_excel(writer, sheet_name="销售明细", index=False)
    region_df.to_excel(writer, sheet_name="地区编码", index=False)
    product_df.to_excel(writer, sheet_name="产品信息", index=False)

print("Excel文件已生成：sales_data.xlsx")