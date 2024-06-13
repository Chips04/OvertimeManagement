import os
import openpyxl
from openpyxl.drawing.image import Image
import pandas as pd


# folder_path: 导出的图片文件夹路径
# file_path： 导出的表格文件路径
# image_file_col： 导出的表格图片所在列名
# col: 列号——处理后的图片在哪列
# output_path: 处理后的文件保存路径
def handle_image(folder_path, file_path, image_file_col, col, output_path):
    # 读取Excel文件
    df = pd.read_excel(file_path, engine='openpyxl')

    # 打开现有的工作簿
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active

    # 遍历表格中的每一行
    for i, row in enumerate(df.itertuples()):
        # 获取当前行的图片文件名
        image_file = getattr(row, image_file_col)
        if image_file:
            # 获取图片文件的全路径
            image_path = os.path.join(folder_path, image_file)
            # 检查文件是否存在
            if os.path.exists(image_path):
                # 将图片嵌入到对应单元格中
                img = Image(image_path)
                img.width = 292  # 设置图片宽度为200像素
                img.height = 220  # 设置图片高度为150像素
                ws.add_image(img, f'{col}{i + 2}')
            else:
                print(f"图片文件 {image_file} 不存在！")
    # 保存修改后的工作簿到新的Excel文件
    wb.save(output_path)


# 执行程序
handle_image(
    "/home/sf107/桌面/上芬社区（包括党委、工作站、党群中心、居委会）获得荣誉汇总登记表_20240604110743/Files/佐证材料",
    "/home/sf107/桌面/上芬社区（包括党委、工作站、党群中心、居委会）获得荣誉汇总登记表_20240604110743/上芬社区（包括党委、工作站、党群中心、居委会）获得荣誉汇总登记表_20240604110743.xlsx",
    '佐证材料',
    'I',
    "/home/sf107/桌面/output_excel_file2.xlsx"
)