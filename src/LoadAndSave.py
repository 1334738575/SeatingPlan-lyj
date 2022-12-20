import os
import shutil
import xlwt
import yaml

# 用于生成了字典:  patterns_dict
color_text = [['', 23],  # 0
              ['领导嘉宾', 23],  # 1
              ['能源', 5],  # 2
              ['公卫', 6],  # 3
              ['航空', 7],  # 4
              ['海外', 24],  # 5
              ['国际', 25],  # 6
              ['化学', 26],  # 7
              ['药学', 27],  # 8
              ['微纳', 29],  # 9
              ['电子', 31],  # 10
              ['生科', 42],  # 11
              ['医学', 44],  # 12
              ['信息', 46],  # 13
              ['环生', 47],  # 14
              ['海地', 49],  # 15
              ['少民', 50],  # 16
              ['隔空', 1]]  # 17

# 用于写入列表头
arab_to_ch = {1: '一', 2: '二', 3: '三', 4: '四', 5: '五',
              6: '六', 7: '七', 8: '八', 9: '九', 10: '十',
              11: '十一', 12: '十二', 13: '十三', 14: '十四', 15: '十五',
              16: '十六', 17: '十七', 18: '十八', 19: '十九', 20: '二十',
              21: '二十一'}


def LoadFromYaml():
    with open("Setting/last_config.yaml", encoding="utf-8") as f:
        mid_info = f.read()
        info = yaml.load(mid_info, Loader=yaml.FullLoader)
    return info


def SaveToYaml(save_to_yaml):
    with open("Setting/last_config.yaml", 'w', encoding="utf-8") as f:
        yaml.dump(save_to_yaml, f, allow_unicode=True)


def Get_Setting(s_path, mod):
    """
    :param s_path: path of setting file
    :param mod: 'DEFAULT' or 'LAST'
    :return: a list of '学院+人数’
    """
    setting_list = []
    setting_path = s_path
    if mod == 'DEFAULT':
        # setting_path = s_path + '/default.txt'
        setting_path = os.path.join(s_path, 'default.txt')
    elif mod == 'LAST':
        # setting_path = s_path + '/last.txt'
        setting_path = os.path.join(s_path, 'last.txt')
    else:
        print('The mode is not unrecognized!')
    setting_file = open(setting_path, encoding='utf-8')
    read_lines = setting_file.readlines()
    for line in read_lines:
        parts = line.split(',')
        setting_list.append([parts[0], int(parts[1])])
    setting_file.close()
    return setting_list


def Save_Setting(s_path, final_dis_list):
    # save_path = s_path + '/last.txt'
    save_path = os.path.join(s_path, 'last.txt')
    save_file = open(save_path, mode=''
                                     'w', encoding='utf-8')
    for i in range(0, len(final_dis_list)):
        save_file.write(final_dis_list[i][0] + ',' + str(final_dis_list[i][1]) + '\n')
    save_file.close()


def CreateNewFile(src, dst):
    shutil.copyfile(src, dst)


def create_xls():
    workbook = xlwt.Workbook(encoding='utf-8')
    sheet = workbook.add_sheet('sheet1')
    return workbook, sheet


def GenerateXlsxSmall(setting_init_path, activity_name, blocks_dict, bool_spaced):
    workbook, sheet = create_xls()
    # 基本属性
    sheet_max_rows = 16
    sheet_max_cols = 34
    # 行高列宽----------------------------------------
    # 格式：sheet.col(n).width = 11 * 256 ，表示第n列的宽度为11个字符
    sheet.col(0).width = 18 * 256  # 实际转换关系还不清楚，第一列19个字符
    for i in range(1, sheet_max_cols):
        if (i == 7) | (i == 26):
            sheet.col(i).width = 6 * 256
        elif (i == 1) | (i == 25) | (i == 32):
            sheet.col(i).width = 6 * 256
        else:
            sheet.col(i).width = 7 * 256
    # 设置高度
    sheet.row(0).height_mismatch = True
    sheet.row(0).height = 900  # 实际值为height/20（磅）
    sheet.row(1).height_mismatch = True
    sheet.row(1).height = 750
    for i in range(2, sheet_max_rows):
        sheet.row(i).height_mismatch = True
        sheet.row(i).height = 525

    # 设置各种字体格式----------------------------
    fonts_dict = {}
    # 标题
    title_font = xlwt.Font()
    title_font.name = '宋体'
    title_font.height = 20 * 28  # 一个20为单位，另一个为实际字号
    title_font.bold = True
    fonts_dict['标题'] = title_font
    # 列表头1，第一列
    col_font = xlwt.Font()
    col_font.name = '宋体'
    col_font.height = 20 * 14  # 一个20为单位，另一个为实际字号
    col_font.bold = True
    fonts_dict['列表头1'] = col_font
    # 列表头2，适用于第二列、第二十六列、以及第33列
    col_font2 = xlwt.Font()
    col_font2.name = 'Times New Roman'
    col_font2.height = 20 * 14  # 一个20为单位，另一个为实际字号
    col_font2.bold = True
    fonts_dict['列表头2'] = col_font2
    # 过道
    aisle_font = xlwt.Font()
    aisle_font.name = 'Times New Roman'
    aisle_font.height = 20 * 18  # 一个20为单位，另一个为实际字号
    aisle_font.bold = True
    fonts_dict['过道'] = aisle_font
    # 空的
    empty_font = xlwt.Font()
    empty_font.name = '宋体'
    empty_font.height = 20 * 24  # 一个20为单位，另一个为实际字号
    empty_font.bold = True
    fonts_dict['空的'] = empty_font
    # 嘉宾和学院
    vip_college_font = xlwt.Font()
    vip_college_font.name = '宋体'
    vip_college_font.height = 20 * 20  # 一个20为单位，另一个为实际字号
    vip_college_font.bold = True
    fonts_dict['嘉宾和学院'] = vip_college_font
    # 主持台
    screen_font = xlwt.Font()
    screen_font.name = '宋体'
    screen_font.height = 20 * 18  # 一个20为单位，另一个为实际字号
    screen_font.bold = True
    fonts_dict['主持台'] = screen_font
    # 总座位数
    remand_font = xlwt.Font()
    remand_font.name = '宋体'
    remand_font.height = 20 * 16  # 一个20为单位，另一个为实际字号
    remand_font.bold = True
    fonts_dict['总座位数'] = remand_font

    # 设置边框格式--------------------------
    # 细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7
    # 大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
    # 一般边框
    general_borders = xlwt.Borders()
    general_borders.left = 1
    general_borders.right = 1
    general_borders.top = 1
    general_borders.bottom = 1

    # 设置对齐方式------------------------------
    align = xlwt.Alignment()
    # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
    align.horz = 0x02
    # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
    align.vert = 0x01

    aisle_align = xlwt.Alignment()
    # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
    aisle_align.horz = 0x02
    # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
    aisle_align.vert = 0x01
    aisle_align.wrap = 1  # 自动换行

    # 设置填充颜色-----------------------------
    default_pattern = xlwt.Pattern()  # Create the Pattern
    default_pattern.pattern = xlwt.Pattern.SOLID_PATTERN  # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
    default_pattern.pattern_fore_colour = 1  # 白色
    patterns_dict = {}
    for i in range(len(color_text)):
        temp_pattern = xlwt.Pattern()
        temp_pattern.pattern = xlwt.Pattern.SOLID_PATTERN
        temp_pattern.pattern_fore_colour = color_text[i][1]
        patterns_dict[color_text[i][0]] = temp_pattern
    # May be: 8 through 63. 0 = Black, 1 = White, 2 = Red, 3 = Green, 4 = Blue, 5 = Yellow, 6 = Magenta, 7 = Cyan,
    # 16 = Maroon, 17 = Dark Green, 18 = Dark Blue, 19 = Dark Yellow , almost brown), 20 = Dark Magenta, 21 = Teal,
    # 22 = Light Gray, 23 = Dark Gray, the list goes on...

    # 写入-------------------------------------------
    # 标题
    title_style = xlwt.XFStyle()
    title_style.font = fonts_dict['标题']
    title_style.alignment = align
    # 合并单元格（r1，r2，c1，c2，v）
    sheet.write_merge(0, 0, 0, sheet_max_cols - 1, activity_name, style=title_style)
    # 行表头
    empty_style = xlwt.XFStyle()
    empty_style.font = fonts_dict['空的']
    empty_style.alignment = align
    sheet.write_merge(1, 1, 2, 6, '空的', style=empty_style)
    sheet.write_merge(1, 1, 27, 31, '空的', style=empty_style)
    # 列表头1
    col_style = xlwt.XFStyle()
    col_style.font = fonts_dict['列表头1']
    col_style.alignment = align
    col_style.borders = general_borders
    for i in range(12):
        sheet.write(i + 1, 0, '第' + arab_to_ch[12 - i] + '排', style=col_style)
    # 列表头2
    col_style2 = xlwt.XFStyle()
    col_style2.font = fonts_dict['列表头2']
    col_style2.alignment = align
    for i in range(10):
        sheet.write(i + 2, 1, 5, style=col_style2)
    sheet.write(12, 1, 3, style=col_style2)
    for i in range(12):
        sheet.write(i + 1, 25, 17, style=col_style2)
    for i in range(11):
        sheet.write(i + 2, 32, 5, style=col_style2)
    # 过道
    aisle_style = xlwt.XFStyle()
    aisle_style.font = fonts_dict['过道']
    aisle_style.alignment = aisle_align
    aisle_style.borders = general_borders
    # aisle_style.alignment.wrap = 1  # 自动换行
    sheet.write_merge(1, 12, 7, 7, '过道', style=aisle_style)
    sheet.write_merge(1, 12, 26, 26, '过道', style=aisle_style)
    # 主持台
    screen_style = xlwt.XFStyle()
    screen_style.font = fonts_dict['主持台']
    screen_style.alignment = align
    screen_style.borders.right = 2
    screen_style.borders.left = 2
    screen_style.borders.top = 2
    screen_style.borders.bottom = 2
    sheet.write_merge(14, 14, 28, 30, '主持台', style=screen_style)

    screen_style2 = xlwt.XFStyle()
    screen_style2.font = fonts_dict['主持台']
    screen_style2.alignment = align
    screen_style2.borders = general_borders
    pattern2 = xlwt.Pattern()
    pattern2.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern2.pattern_fore_colour = 22
    screen_style2.pattern = pattern2
    sheet.write_merge(14, 14, 4, 27, style=screen_style2)
    # 总座位数
    remand_style = xlwt.XFStyle()
    remand_style.font = fonts_dict['总座位数']
    remand_style.alignment = align
    sheet.write_merge(15, 15, 0, 1, '总座位数：312', style=remand_style)
    # 嘉宾和学院 TODO
    text_vip_college_style = xlwt.XFStyle()  # 这里不能写在for循环的外面，不然会共享一个style
    text_vip_college_style.font = fonts_dict['嘉宾和学院']
    text_vip_college_style.alignment = align
    # 记录平均坐标以及座位数
    det_row = 1
    max_row = 12
    for [block, det_col, max_col] in [['l', 2, 5], ['m', 8, 17], ['r', 27, 5]]:
        write_to_left = blocks_dict[block]
        # 平均坐标以及座位数存储list
        write_text = [[0 for i in range(4)] for j in range(17)]
        for i in range(max_row):
            for j in range(max_col):
                if LocNotWriteInSmall(i, j, block):
                    continue
                ii = i + det_row
                jj = j + det_col
                if bool_spaced:
                    if (j % 2) == 0:
                        write_text[write_to_left[i][j]][0] += 1
                else:
                    write_text[write_to_left[i][j]][0] += 1
                write_text[write_to_left[i][j]][3] += 1
                write_text[write_to_left[i][j]][1] += ii
                write_text[write_to_left[i][j]][2] += jj
        for i in range(17):
            if write_text[i][0] == 0:
                continue
            write_text[i][1] = int(round(write_text[i][1] / write_text[i][3]))
            if bool_spaced:
                write_text[i][2] = int(write_text[i][2] / write_text[i][3]) - 1  # TODO
            else:
                write_text[i][2] = int(write_text[i][2] / write_text[i][3])  # TODO
            # str_num = str(write_text[i][0])
            # sheet.write_merge(write_text[i][1], write_text[i][1], write_text[i][2], write_text[i][2]+1,
            #                   color_text[i][0]+str_num, style=text_vip_college_style)
        # print('第一次写：', write_text[0][1], write_text[0][2])
        # 只需要判断第一行的上边界，以及最第一列的左边界
        next_continue = False
        first_row_continue = 0
        for i in range(max_row):
            for j in range(max_col):
                if LocNotWriteInSmall(i, j, block):
                    continue
                if next_continue:
                    next_continue = False
                    # print('jump')
                    continue
                if first_row_continue != 0:
                    first_row_continue -= 1
                    continue
                vip_college_style = xlwt.XFStyle()  # 这里不能写在for循环的外面，不然会共享一个style
                vip_college_style.font = fonts_dict['嘉宾和学院']
                vip_college_style.alignment = align
                vip_college_pattern = xlwt.Pattern()
                vip_college_pattern.pattern = xlwt.Pattern.SOLID_PATTERN
                borders = xlwt.Borders()
                # 这里逻辑有漏洞，但是不影响正常使用
                # 下
                if block == 'l':
                    if (i == max_row - 2) & (write_to_left[i][j] != 0):
                        if (j == 0) | (j == 1):
                            borders.bottom = 5
                if i == (max_row - 1):
                    borders.bottom = 5
                elif write_to_left[i][j] != write_to_left[i + 1][j]:
                    borders.bottom = 5
                # 上
                if (block == 'l') | (block == 'r'):
                    if (i == 1) & (write_to_left[i][j] != 0):
                        borders.top = 5
                elif (i == 0) & (write_to_left[i][j] != 0):
                    borders.top = 5
                # 左
                if (block == 'l') & (i == max_row - 1) & (j == 2):
                    borders.left = 5
                elif (j == 0) & (write_to_left[i][j] != 0):
                    borders.left = 5
                # 右
                if j == (max_col - 1):
                    if write_to_left[i][j] != 0:
                        borders.right = 5
                elif write_to_left[i][j] != write_to_left[i][j + 1]:
                    borders.right = 5
                vip_college_style.borders = borders
                if bool_spaced & ((j % 2) == 1) & (write_to_left[i][j] != 0) \
                        & (write_to_left[i][j] != 1) & (write_to_left[i][j] != 17):
                    vip_college_pattern.pattern_fore_colour = 1  # 隔座时，隔列填白 TODO
                else:
                    vip_college_pattern.pattern_fore_colour = color_text[write_to_left[i][j]][1]
                vip_college_style.pattern = vip_college_pattern
                ii = i + det_row
                jj = j + det_col
                if (block == 'l') & (write_to_left[i][j] == 1) & (ii == (max_row - 1 + det_row)):
                    idx = 1
                    vip_college_style.borders.left = 5
                    vip_college_style.borders.right = 5
                    # if
                    sheet.write_merge(ii, ii, 4,
                                      6,
                                      color_text[idx][0], style=vip_college_style)
                    first_row_continue = 2
                    continue
                elif (write_to_left[i][j] == 1) & (ii == (max_row - 1 + det_row)) \
                        & (jj == (write_text[1][2])):
                    idx = 1
                    sheet.write_merge(ii, ii, write_text[idx][2],
                                      write_text[idx][2] + 2,
                                      color_text[idx][0], style=vip_college_style)
                    first_row_continue = 2
                    continue
                if bool_spaced:
                    idx = write_to_left[i][j]
                    if (ii == write_text[write_to_left[i][j]][1]) & (jj == write_text[write_to_left[i][j]][2]) \
                            & (write_to_left[i][j] != 0) & (write_to_left[i][j] != 1):
                        sheet.write(write_text[idx][1], write_text[idx][2],
                                    color_text[idx][0][0], style=vip_college_style)
                    elif (ii == write_text[write_to_left[i][j]][1]) & (jj == write_text[write_to_left[i][j]][2]+1) \
                            & (write_to_left[i][j] != 0) & (write_to_left[i][j] != 1):
                        sheet.write(write_text[idx][1], write_text[idx][2] + 1,
                                    color_text[idx][0][1], style=vip_college_style)
                    elif (ii == write_text[write_to_left[i][j]][1]) & (jj == write_text[write_to_left[i][j]][2]+2) \
                            & (write_to_left[i][j] != 0) & (write_to_left[i][j] != 1):
                        int_num = write_text[write_to_left[i][j]][0]
                        sheet.write(write_text[idx][1], write_text[idx][2] + 2,
                                    int_num, style=vip_college_style)
                    else:
                        sheet.write(ii, jj, style=vip_college_style)
                else:
                    if (ii == write_text[write_to_left[i][j]][1]) & (jj == write_text[write_to_left[i][j]][2]) \
                            & (write_to_left[i][j] != 0):
                        # print('发生：', ii, jj)
                        idx = write_to_left[i][j]
                        str_num = str(write_text[write_to_left[i][j]][0])
                        if (write_text[write_to_left[i][j]][0] == 2) | (write_text[write_to_left[i][j]][0] == 3):
                            vip_college_style.borders.right = 5
                        sheet.write_merge(write_text[idx][1], write_text[idx][1], write_text[idx][2],
                                          write_text[idx][2] + 1,
                                          color_text[idx][0] + str_num, style=vip_college_style)
                        next_continue = True
                    else:
                        sheet.write(ii, jj, style=vip_college_style)

    workbook.save(setting_init_path)


def LocNotWriteInSmall(i, j, block):
    if (block != 'm') & (i == 0):
        return True
    if (block == 'l') & (i == 11):
        if (j == 0) | (j == 1):
            return True
    return False


def GenerateXlsxBig(setting_init_path, activity_name, blocks_dict, bool_spaced):
    workbook, sheet = create_xls()
    # 基本属性
    sheet_max_rows = 27
    sheet_max_cols = 33
    # 行高列宽----------------------------------------
    # 格式：sheet.col(n).width = 11 * 256 ，表示第n列的宽度为11个字符
    sheet.col(0).width = 19 * 256  # 实际转换关系还不清楚，第一列19个字符
    for i in range(1, sheet_max_cols):
        sheet.col(i).width = 7 * 256
    # 设置高度
    sheet.row(0).height_mismatch = True
    sheet.row(0).height = 750  # 实际值为height/20（磅）
    sheet.row(1).height_mismatch = True
    sheet.row(1).height = 750
    for i in range(2, sheet_max_rows):
        sheet.row(i).height_mismatch = True
        sheet.row(i).height = 525

    # 设置各种字体格式----------------------------
    fonts_dict = {}
    # 标题
    title_font = xlwt.Font()
    title_font.name = '宋体'
    title_font.height = 20 * 24  # 一个20为单位，另一个为实际字号
    title_font.bold = True
    fonts_dict['标题'] = title_font
    # 列表头
    col_font = xlwt.Font()
    col_font.name = '宋体'
    col_font.height = 20 * 14  # 一个20为单位，另一个为实际字号
    col_font.bold = True
    fonts_dict['列表头'] = col_font
    # 行表头
    row_font = xlwt.Font()
    row_font.name = 'Times New Roman'
    row_font.height = 20 * 24  # 一个20为单位，另一个为实际字号
    row_font.bold = True
    fonts_dict['行表头'] = row_font
    # 嘉宾和学院
    vip_college_font = xlwt.Font()
    vip_college_font.name = '宋体'
    vip_college_font.height = 20 * 20  # 一个20为单位，另一个为实际字号
    vip_college_font.bold = True
    fonts_dict['嘉宾和学院'] = vip_college_font
    # 银幕
    screen_font = xlwt.Font()
    screen_font.name = '宋体'
    screen_font.height = 20 * 18  # 一个20为单位，另一个为实际字号
    screen_font.bold = True
    fonts_dict['银幕'] = screen_font
    # 总座位数
    remand_font = xlwt.Font()
    remand_font.name = '宋体'
    remand_font.height = 20 * 16  # 一个20为单位，另一个为实际字号
    remand_font.bold = True
    fonts_dict['总座位数'] = remand_font

    # 设置边框格式--------------------------
    # 细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7
    # 大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
    # 一般边框
    general_borders = xlwt.Borders()
    general_borders.left = 1
    general_borders.right = 1
    general_borders.top = 1
    general_borders.bottom = 1

    # 设置对齐方式------------------------------
    align = xlwt.Alignment()
    # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
    align.horz = 0x02
    # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
    align.vert = 0x01

    # 设置填充颜色-----------------------------
    default_pattern = xlwt.Pattern()  # Create the Pattern
    default_pattern.pattern = xlwt.Pattern.SOLID_PATTERN  # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
    default_pattern.pattern_fore_colour = 1  # 白色
    patterns_dict = {}
    for i in range(len(color_text)):
        temp_pattern = xlwt.Pattern()
        temp_pattern.pattern = xlwt.Pattern.SOLID_PATTERN
        temp_pattern.pattern_fore_colour = color_text[i][1]
        patterns_dict[color_text[i][0]] = temp_pattern
    # May be: 8 through 63. 0 = Black, 1 = White, 2 = Red, 3 = Green, 4 = Blue, 5 = Yellow, 6 = Magenta, 7 = Cyan,
    # 16 = Maroon, 17 = Dark Green, 18 = Dark Blue, 19 = Dark Yellow , almost brown), 20 = Dark Magenta, 21 = Teal,
    # 22 = Light Gray, 23 = Dark Gray, the list goes on...

    # 写入-------------------------------------------
    # 标题
    title_style = xlwt.XFStyle()
    title_style.font = fonts_dict['标题']
    title_style.alignment = align
    # 合并单元格（r1，r2，c1，c2，v）
    sheet.write_merge(0, 0, 0, sheet_max_cols - 1, activity_name, style=title_style)
    # 行表头
    row_style = xlwt.XFStyle()
    row_style.font = fonts_dict['行表头']
    row_style.alignment = align
    for i in range(5):
        sheet.write(1, 2 + i, int(i + 1), style=row_style)
    for i in range(17):
        sheet.write(1, 9 + i, int(i + 6), style=row_style)
    for i in range(5):
        sheet.write(1, 28 + i, int(i + 23), style=row_style)
    # 列表头
    col_style = xlwt.XFStyle()
    col_style.font = fonts_dict['列表头']
    col_style.alignment = align
    col_style.borders = general_borders
    for i in range(21):
        sheet.write(i + 2, 0, '第' + arab_to_ch[21 - i] + '排', style=col_style)
    # 银幕
    screen_style = xlwt.XFStyle()
    screen_style.font = fonts_dict['银幕']
    screen_style.alignment = align
    screen_style.borders = general_borders
    pattern = xlwt.Pattern()
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern.pattern_fore_colour = 22
    screen_style.pattern = pattern
    sheet.write_merge(24, 24, 4, 28, '银幕', style=screen_style)
    # 总座位数
    remand_style = xlwt.XFStyle()
    remand_style.font = fonts_dict['总座位数']
    remand_style.alignment = align
    sheet.write_merge(26, 26, 0, 5, '总座位数：563', style=remand_style)
    # 嘉宾和学院 TODO
    text_vip_college_style = xlwt.XFStyle()  # 这里不能写在for循环的外面，不然会共享一个style
    text_vip_college_style.font = fonts_dict['嘉宾和学院']
    text_vip_college_style.alignment = align
    # 记录平均坐标以及座位数
    det_row = 2
    max_row = 21
    for [block, det_col, max_col] in [['l', 2, 5], ['m', 9, 17], ['r', 28, 5]]:
        write_to_left = blocks_dict[block]  # 实际为写入每个板块的list
        write_text = [[0 for i in range(4)] for j in range(17)]
        for i in range(max_row):
            for j in range(max_col):
                ii = i + det_row
                jj = j + det_col
                if bool_spaced:
                    if (j % 2) == 0:
                        write_text[write_to_left[i][j]][0] += 1
                else:
                    write_text[write_to_left[i][j]][0] += 1
                write_text[write_to_left[i][j]][3] += 1
                write_text[write_to_left[i][j]][1] += ii
                write_text[write_to_left[i][j]][2] += jj
        for i in range(17):
            if write_text[i][0] == 0:
                continue
            write_text[i][1] = int(round(write_text[i][1] / write_text[i][3]))
            if bool_spaced:
                write_text[i][2] = int(write_text[i][2] / write_text[i][3]) - 1  # TODO
            else:
                write_text[i][2] = int(write_text[i][2] / write_text[i][3])  # TODO
            # str_num = str(write_text[i][0])
            # sheet.write_merge(write_text[i][1], write_text[i][1], write_text[i][2], write_text[i][2]+1,
            #                   color_text[i][0]+str_num, style=text_vip_college_style)
        # print('第一次写：', write_text[0][1], write_text[0][2])
        # 只需要判断第一行的上边界，以及最第一列的左边界
        next_continue = False
        first_row_continue = 0
        for i in range(max_row):
            for j in range(max_col):
                if next_continue:
                    next_continue = False
                    # print('jump')
                    continue
                if first_row_continue != 0:
                    first_row_continue -= 1
                    continue
                vip_college_style = xlwt.XFStyle()  # 这里不能写在for循环的外面，不然会共享一个style
                vip_college_style.font = fonts_dict['嘉宾和学院']
                vip_college_style.alignment = align
                vip_college_pattern = xlwt.Pattern()
                vip_college_pattern.pattern = xlwt.Pattern.SOLID_PATTERN
                borders = xlwt.Borders()
                # 这里逻辑有漏洞，但是不影响正常使用
                # 下
                if i == (max_row - 1):
                    borders.bottom = 5
                elif write_to_left[i][j] != write_to_left[i + 1][j]:
                    borders.bottom = 5
                # 上
                if (i == 0) & (write_to_left[i][j] != 0):
                    borders.top = 5
                # 左
                if (j == 0) & (write_to_left[i][j] != 0):
                    borders.left = 5
                # 右
                if j == (max_col - 1):
                    if write_to_left[i][j] != 0:
                        borders.right = 5
                elif write_to_left[i][j] != write_to_left[i][j + 1]:
                    borders.right = 5
                vip_college_style.borders = borders
                if bool_spaced & ((j % 2) == 1) & (write_to_left[i][j] != 0) \
                        & (write_to_left[i][j] != 1) & (write_to_left[i][j] != 17):
                    vip_college_pattern.pattern_fore_colour = 1  # 隔座时，隔列填白 TODO
                else:
                    vip_college_pattern.pattern_fore_colour = color_text[write_to_left[i][j]][1]
                vip_college_style.pattern = vip_college_pattern
                ii = i + det_row
                jj = j + det_col
                if (write_to_left[i][j] == 1) & (ii == write_text[1][1]) & (jj == (write_text[1][2])):
                    idx = 1
                    sheet.write_merge(write_text[idx][1], write_text[idx][1], write_text[idx][2],
                                      write_text[idx][2] + 2,
                                      color_text[idx][0], style=vip_college_style)
                    first_row_continue = 2
                    continue
                if bool_spaced:
                    idx = write_to_left[i][j]
                    if (ii == write_text[write_to_left[i][j]][1]) & (jj == write_text[write_to_left[i][j]][2]) \
                            & (write_to_left[i][j] != 0) & (write_to_left[i][j] != 1):
                        sheet.write(write_text[idx][1], write_text[idx][2],
                                    color_text[idx][0][0], style=vip_college_style)
                    elif (ii == write_text[write_to_left[i][j]][1]) & (jj == write_text[write_to_left[i][j]][2]+1) \
                            & (write_to_left[i][j] != 0) & (write_to_left[i][j] != 1):
                        sheet.write(write_text[idx][1], write_text[idx][2] + 1,
                                    color_text[idx][0][1], style=vip_college_style)
                    elif (ii == write_text[write_to_left[i][j]][1]) & (jj == write_text[write_to_left[i][j]][2]+2) \
                            & (write_to_left[i][j] != 0) & (write_to_left[i][j] != 1):
                        int_num = write_text[write_to_left[i][j]][0]
                        sheet.write(write_text[idx][1], write_text[idx][2] + 2,
                                    int_num, style=vip_college_style)
                    else:
                        sheet.write(ii, jj, style=vip_college_style)
                else:
                    if (ii == write_text[write_to_left[i][j]][1]) & (jj == write_text[write_to_left[i][j]][2]) \
                            & (write_to_left[i][j] != 0):
                        # print('发生：', ii, jj)
                        idx = write_to_left[i][j]
                        str_num = str(write_text[write_to_left[i][j]][0])
                        if (write_text[write_to_left[i][j]][0] == 2) | (write_text[write_to_left[i][j]][0] == 3):
                            vip_college_style.borders.right = 5
                        sheet.write_merge(write_text[idx][1], write_text[idx][1], write_text[idx][2],
                                          write_text[idx][2] + 1,
                                          color_text[idx][0] + str_num, style=vip_college_style)
                        next_continue = True
                    else:
                        sheet.write(ii, jj, style=vip_college_style)

    # 过道
    sheet.write_merge(1, 22, 7, 8)
    sheet.write_merge(1, 22, 26, 27)

    workbook.save(setting_init_path)


if __name__ == '__main__':
    pwd_path = os.getcwd()
    test_path = os.path.join(pwd_path, 'Result', 'test.xlsx')
    print(test_path)
    act_name = 'test'
    dict1 = {}
    # GenerateXlsxBig(test_path, act_name, dict1)
    GenerateXlsxSmall(test_path, act_name, dict1)

    ##############
    # str1 = 'abcd'
    # print(str1[1])
