import os
import subprocess
import webbrowser
from invoice_template_handler import save_header_template, save_footer_template, generate_invoice

# 实际业务中的头部模板
real_header = """
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>商业发票</title>
    <style>
        table { width: 100%; border-collapse: collapse; }
        th, td { border: 1px solid black; padding: 5px; }
        th { background-color: #f2f2f2; }
    </style>
</head>
<body>
    <table>
        <tr>
            <th colspan="8">Shibo Chuangxiang Digital Technology (Shenzhen) Co., LTD</th>
        </tr>
        <tr>
            <td colspan="8">Room 1501, Shenzhen International Qianhai Yidu Tower, No.99, Gangcheng Street, Nanshan Street, Qianhai Shenzhen-Hong Kong Cooperation Zone, Shenzhen.</td>
        </tr>
        <tr>
            <th colspan="8">Commercial Invoice</th>
        </tr>
        <tr>
            <td>Buyer: UNICAIR(HOLDINGS) LIMITED</td>
            <td colspan="3"></td>
            <td>CI No.:</td>
            <td>CXCI2025012201</td>
        </tr>
        <tr>
            <td>ADD:UNIT 802,8/F,CHINA INSURANCE GROUP</td>
            <td colspan="3"></td>
            <td>Date:</td>
            <td>2025/1/22</td>
        </tr>
        <tr>
            <td>BUILDING,141 DES VOEUX ROAD CENTRAL HONG KONG</td>
            <td colspan="3"></td>
            <td>PO No.:</td>
            <td></td>
        </tr>
        <tr>
            <th>NO.</th>
            <th>Material code</th>
            <th>DESCRIPTION</th>
            <th>Model NO.</th>
            <th>Unit Price USD</th>
            <th>Qty</th>
            <th>Unit</th>
            <th>Amount USD</th>
        </tr>
"""

# 实际业务中的尾部模板
real_footer = """
        <tr>
            <td></td>
            <td>TTL:</td>
            <td></td>
            <td></td>
            <td></td>
            <td>9506</td>
            <td></td>
            <td>29,847.38</td>
        </tr>
    </table>
    <p>SAY USD TWENTY-NINE THOUSAND EIGHT HUNDRED AND FORTY-SEVEN AND POINT THIRTY-EIGHT ONLY.</p>
    <p>COUNTRY OF ORIGIN: CHINA</p>
    <p>Payment Term:100% TT within 5 working days when Unicair(Holdings) Limited receive Goods.</p>
    <p>Delivery Term:CIF</p>
    <p>Company Name：Shibo Chuangxiang Digital Technology (Shenzhen) Co., LTD</p>
    <p>Account number：811030101280058437</p>
    <p>Bank Name: China citic bank shenzhen branch</p>
    <p>Bank Address:8F,Citic security tower, zhongxin 4road, futian dist. futian shenzhen china</p>
    <p>SWIFT No.: CIBKCNBJ518</p>
</body>
</html>
"""

# 从CSV或Excel读取的商品数据（模拟）
def get_items_from_data(data):
    """从数据生成商品列表HTML"""
    items_html = ""
    for i, item in enumerate(data, 1):
        code, desc, model, price, qty, unit, amount = item
        items_html += f"""
        <tr>
            <td>{i}</td>
            <td>{code}</td>
            <td>{desc}</td>
            <td>{model}</td>
            <td>{price}</td>
            <td>{qty}</td>
            <td>{unit}</td>
            <td>{amount}</td>
        </tr>
        """
    return items_html

# 模拟数据
sample_data = [
    ("C100.C05-032-04-00", "铣刀", "/", 0.7318, 500, "个", 365.90),
    ("E100.020310008", "红外发热管", "/", 16.1939, 6, "个", 97.16),
    ("E100.020310014", "联轴器", "12-14", 11.4651, 2, "个", 22.93),
    ("E100.020396056", "发热板", "/", 48.4829, 2, "个", 96.97)
]

# 保存模板
save_header_template(real_header)
save_footer_template(real_footer)

# 生成商品列表
items_content = get_items_from_data(sample_data)

# 在浏览器中打开查看效果

您可以通过以下几种方式在浏览器中打开生成的HTML文件查看效果:

### 方法一：手动打开

1. 打开文件资源管理器，导航到 `D:\project\declare_list` 目录
2. 找到 `business_invoice.html` 文件
3. 双击该文件，它会在默认浏览器中打开

### 方法二：修改代码自动打开

您可以修改 `business_test.py` 文件，添加自动打开浏览器的功能：
```python
import os
import webbrowser
from invoice_template_handler import save_header_template, save_footer_template, generate_invoice

# 实际业务中的头部模板
real_header = """
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>商业发票</title>
    <style>
        table { width: 100%; border-collapse: collapse; }
        th, td { border: 1px solid black; padding: 5px; }
        th { background-color: #f2f2f2; }
    </style>
</head>
<body>
    <table>
        <tr>
            <th colspan="8">Shibo Chuangxiang Digital Technology (Shenzhen) Co., LTD</th>
        </tr>
        <tr>
            <td colspan="8">Room 1501, Shenzhen International Qianhai Yidu Tower, No.99, Gangcheng Street, Nanshan Street, Qianhai Shenzhen-Hong Kong Cooperation Zone, Shenzhen.</td>
        </tr>
        <tr>
            <th colspan="8">Commercial Invoice</th>
        </tr>
        <tr>
            <td>Buyer: UNICAIR(HOLDINGS) LIMITED</td>
            <td colspan="3"></td>
            <td>CI No.:</td>
            <td>CXCI2025012201</td>
        </tr>
        <tr>
            <td>ADD:UNIT 802,8/F,CHINA INSURANCE GROUP</td>
            <td colspan="3"></td>
            <td>Date:</td>
            <td>2025/1/22</td>
        </tr>
        <tr>
            <td>BUILDING,141 DES VOEUX ROAD CENTRAL HONG KONG</td>
            <td colspan="3"></td>
            <td>PO No.:</td>
            <td></td>
        </tr>
        <tr>
            <th>NO.</th>
            <th>Material code</th>
            <th>DESCRIPTION</th>
            <th>Model NO.</th>
            <th>Unit Price USD</th>
            <th>Qty</th>
            <th>Unit</th>
            <th>Amount USD</th>
        </tr>
"""

# 实际业务中的尾部模板
real_footer = """
        <tr>
            <td></td>
            <td>TTL:</td>
            <td></td>
            <td></td>
            <td></td>
            <td>9506</td>
            <td></td>
            <td>29,847.38</td>
        </tr>
    </table>
    <p>SAY USD TWENTY-NINE THOUSAND EIGHT HUNDRED AND FORTY-SEVEN AND POINT THIRTY-EIGHT ONLY.</p>
    <p>COUNTRY OF ORIGIN: CHINA</p>
    <p>Payment Term:100% TT within 5 working days when Unicair(Holdings) Limited receive Goods.</p>
    <p>Delivery Term:CIF</p>
    <p>Company Name：Shibo Chuangxiang Digital Technology (Shenzhen) Co., LTD</p>
    <p>Account number：811030101280058437</p>
    <p>Bank Name: China citic bank shenzhen branch</p>
    <p>Bank Address:8F,Citic security tower, zhongxin 4road, futian dist. futian shenzhen china</p>
    <p>SWIFT No.: CIBKCNBJ518</p>
</body>
</html>
"""

# 从CSV或Excel读取的商品数据（模拟）
def get_items_from_data(data):
    """从数据生成商品列表HTML"""
    items_html = ""
    for i, item in enumerate(data, 1):
        code, desc, model, price, qty, unit, amount = item
        items_html += f"""
        <tr>
            <td>{i}</td>
            <td>{code}</td>
            <td>{desc}</td>
            <td>{model}</td>
            <td>{price}</td>
            <td>{qty}</td>
            <td>{unit}</td>
            <td>{amount}</td>
        </tr>
        """
    return items_html

# 模拟数据
sample_data = [
    ("C100.C05-032-04-00", "铣刀", "/", 0.7318, 500, "个", 365.90),
    ("E100.020310008", "红外发热管", "/", 16.1939, 6, "个", 97.16),
    ("E100.020310014", "联轴器", "12-14", 11.4651, 2, "个", 22.93),
    ("E100.020396056", "发热板", "/", 48.4829, 2, "个", 96.97)
]

# 保存模板
save_header_template(real_header)
save_footer_template(real_footer)

# 生成商品列表
items_content = get_items_from_data(sample_data)

# 生成发票
output_path = os.path.join(os.path.dirname(__file__), "business_invoice.html")
generate_invoice(items_content, output_path)

print(f"业务发票已生成: {output_path}")
print("正在打开浏览器查看效果...")

# 使用webbrowser模块自动打开生成的HTML文件
webbrowser.open('file://' + os.path.abspath(output_path))