
# coding: utf-8

def set_style(name, height, color, bold=False):
    style = xlwt.XFStyle()  # 初始化样式

    font = xlwt.Font()  # 为样式创建字体
    # 字体类型：比如宋体、仿宋也可以是汉仪瘦金书繁
    font.name = name
    # 是否为粗体
    font.bold = bold
    # 设置字体颜色
    font.colour_index = color
    # 字体大小
    font.height = height
    # 字体是否斜体
    font.italic = True
    # 字体下划,当值为11时。填充颜色就是蓝色
    font.underline = 0
    # 字体中是否有横线struck_out
    font.struck_out = True
    # 定义格式
    style.font = font

    return style


if __name__ == '__main__':
    import xlwt
    # 创建工作簿,并指定写入的格式
    f = xlwt.Workbook(encoding='utf8')  # 创建工作簿

    #  创建sheet，并指定可以重复写入数据的情况.设置行高度
    sheet1 = f.add_sheet(u'colour', cell_overwrite_ok=False)

    # 控制行的位置
    column = 0
    row = 0
    # 生成第一行
    for i in range(0, 100):
        # 参数对应：行，列，值，字体样式(可以没有)
        sheet1.write(column, row, i, set_style(u'汉仪瘦金书繁', 400, i, False))

        # 这里主要为了控制输入每行十个内容。为了查看
        row = row + 1
        if row % 10 == 0:
            column = column + 1
            row = 0
    f.save("1.xls")