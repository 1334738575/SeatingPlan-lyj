import copy
from LoadAndSave import GenerateXlsxBig, GenerateXlsxSmall

idx_dict = {
    '空': 0,
    '领导嘉宾': 1,
    '能源学院': 2,
    '公共卫生学院': 3,
    '航空航天学院': 4,
    '海外教育学院': 5,
    '国际学院': 6,
    '化学化工学院': 7,
    '药学院': 8,
    '萨本栋微米纳米科学技术研究院': 9,
    '电子科学与技术学院': 10,
    '生命科学学院': 11,
    '医学院': 12,
    '信息学院': 13,
    '环境与生态学院': 14,
    '海洋与地球学院': 15,
    '少数民族预科班': 16,
    '隔空': 1
}

bk_w_1 = 5  # 3~7
bk_w_2 = 17  # 10~26
bk_w_3 = 5  # 29~33
b_bk_w_1 = 5  # 3~7
b_bk_w_2 = 17  # 10~26
b_bk_w_3 = 5  # 29~33
all_num = 0  # 总人数
pre_rows = 0  # 大概需要的排数
det_rows = 1  # 主-侧 容许的排差
max_rows = 21  # 报告厅最大排数
bool_spaced = False  # 是否隔座


def InBigHall(setting_init_path, final_list, vip_rows, activity_name, spaced):
    global max_rows, bool_spaced
    if spaced == 1:
        bool_spaced = True
    else:
        bool_spaced = False
    max_rows = 21
    seating_result = CorrespondingToSeat(final_list, 'Big')
    if pre_rows + vip_rows >= max_rows:
        return False
    result_dict = SortResult(seating_result)
    # print(result_dict)
    # for block in ['l', 'm', 'r']:
    #     for i in range(0, len(result_dict[block])):
    #         print(result_dict[block][i])
    blocks_dict = AssignToGlobal(result_dict, 'Big', vip_rows)
    # print(result_dict)
    # for block in ['l', 'm', 'r']:
    #     for i in range(0, len(blocks_dict[block])):
    #         print(blocks_dict[block][i])
    GenerateXlsxBig(setting_init_path, activity_name, blocks_dict, bool_spaced)
    return True


def InSmallHall(setting_init_path, final_list, vip_rows, activity_name, spaced):
    global max_rows, bool_spaced
    if spaced == 1:
        bool_spaced = True
    else:
        bool_spaced = False
    max_rows = 12
    if vip_rows > 0:
        seating_result = CorrespondingToSeat(final_list, 'Big')
    else:
        seating_result = CorrespondingToSeat(final_list, 'Small')
    if pre_rows + vip_rows >= max_rows:
        return False
    result_dict = SortResult(seating_result)
    # print(result_dict)
    # for block in ['l', 'm', 'r']:
    #     for i in range(0, len(result_dict[block])):
    #         print(result_dict[block][i])
    blocks_dict = AssignToGlobal(result_dict, 'Small', vip_rows)
    # for block in ['l', 'm', 'r']:
    #     for i in range(0, len(blocks_dict[block])):
    #         print(blocks_dict[block][i])
    GenerateXlsxSmall(setting_init_path, activity_name, blocks_dict, bool_spaced)
    return True


def SortResult(seating_result):
    result_dict = {'l': [], 'm': [], 'r': []}
    for i in range(0, len(seating_result)):
        block = seating_result[i][1]
        result_dict[block].append(seating_result[i])
    return result_dict


def AlignInv(align_mod, block):
    if align_mod[block] == 'l':
        align_mod[block] = 'r'
    else:
        align_mod[block] = 'l'


def AlignEveryRows(mod):
    aligns = {'m': [], 'l': [], 'r': []}
    for i in range(0, max_rows):
        aligns['m'].append('l')
        aligns['l'].append('r')
        aligns['r'].append('l')
    return aligns


def AssignToGlobal(result_dict, mod, vip_rows):
    """
    :param result_dict:  list[name, block, start[row, col], end[row, col]]
    :param mod: 'Big' or 'Small'
    :param vip_rows:
    :return: array of every block
    """
    rows_num = {'m': bk_w_2 - 1, 'l': bk_w_1 - 1, 'r': bk_w_3 - 1}
    spaced_rows_num = {'m': bk_w_2 - 1, 'l': bk_w_1 - 1, 'r': bk_w_3 - 1}  # 不考虑隔座影响的每排座位数
    aligns = AlignEveryRows(mod)  # 每行的排列方式
    blocks_dict = {'m': [], 'l': [], 'r': []}
    # 判断排序走向
    for block in ['l', 'm', 'r']:  # 三个板块
        default_align = aligns[block][0]  # 该板块默认排序走向
        seat_result_block = result_dict[block]
        num_college_block = len(result_dict[block])
        rows_num_b = rows_num[block]
        spaced_rows_num_b = spaced_rows_num[block]
        id_inv = -1
        for i in range(0, num_college_block):
            college = seat_result_block[i]
            if bool_spaced:
                college[2][1] *= 2
                college[3][1] *= 2
                if college[3][1] != rows_num_b:
                    college[3][1] += 1
            # 算上领导嘉宾排数
            college[2][0] += vip_rows
            college[3][0] += vip_rows
            # 进行头尾修正
            if (college[3][0] - college[2][0] > 0) & (college[3][1] != rows_num_b) & \
                    (college[3][0] - id_inv > 1):  # 发生跨行，末行未满，离上一次变向超过一行
                AlignInv(aligns[block], college[3][0])  # 该学院的末行反向排序
                id_inv = college[3][0]

        block_seat = [[0 for val in range(spaced_rows_num_b + 1)] for val2 in range(max_rows)]
        for i in range(vip_rows):
            for j in range(spaced_rows_num_b + 1):
                block_seat[i][j] = idx_dict['领导嘉宾']
        for i in range(vip_rows, max_rows):
            for j in range(spaced_rows_num_b + 1):
                idx = GetIdx(i, j, seat_result_block, num_college_block)
                block_seat[i][j] = idx
        for i in range(vip_rows, max_rows):
            if aligns[block][i] != default_align:
                ex_size = int((spaced_rows_num_b + 1) / 2)
                for j in range(ex_size):
                    temp_idx = block_seat[i][j]
                    block_seat[i][j] = block_seat[i][spaced_rows_num_b - j]
                    block_seat[i][spaced_rows_num_b - j] = temp_idx
        # 小报告厅，左板块的左下角两个座位去除
        if (mod == 'Small') & (block == 'l'):
            block_seat[0][0] = -1
            block_seat[0][1] = -1
        # 转换为方便写入的顺序
        down_to_up = []
        for i in range(len(block_seat)):
            down_to_up.append(block_seat[len(block_seat) - 1 - i])
        blocks_dict[block] = down_to_up
    return blocks_dict
    # for block in ['l', 'm', 'r']:
    #     print(block, ': ')
    #     i = 0
    #     for one_row in blocks_dict[block]:
    #         i += 1
    #         print(i, ' ', one_row)


def GetIdx(row, col, seat_result_block, num_college_block):
    for i in range(num_college_block):
        start = seat_result_block[i][2]
        end = seat_result_block[i][3]
        if row < start[0]:
            continue
        elif (row == start[0]) & (col < start[1]):
            continue
        elif row > end[0]:
            continue
        elif (row == end[0]) & (col > end[1]):
            continue
        return idx_dict[seat_result_block[i][0]]
    return idx_dict['空']


def CorrespondingToSeat(final_list, mod):
    """
    :param final_list: 人数分配时，不为0的学院名+对应人数
    :param mod: 'Big' 或者 'Small'
    :return: [学院名 板块号'm''l''r' [起始排号 起始列号] [末尾排号 末尾列号]]
    """
    global pre_rows  # 大概需要的排数
    global b_bk_w_1, b_bk_w_2, b_bk_w_3
    # 因为隔座，每排的座位数减半，若原数目为单数，则+1
    if bool_spaced:
        b_bk_w_1 = int(bk_w_1 / 2) + 1
        b_bk_w_2 = int(bk_w_2 / 2) + 1
        b_bk_w_3 = int(bk_w_3 / 2) + 1
    else:
        b_bk_w_1 = bk_w_1
        b_bk_w_2 = bk_w_2
        b_bk_w_3 = bk_w_3
    # print(b_bk_w_1)
    # print(b_bk_w_2)
    # print(b_bk_w_3)

    s_list = SortList(final_list)  # 将各学院按人数从多到少排列
    num_college = len(s_list)  # 需要排座的学院个数

    seating_result = []  # 最终排座结果：[学院名 板块号'm''l''r' [起始排号 起始列号] [末尾排号 末尾列号]]

    # 各板块的下一个分配开始位置, [行，列]
    m_cur_loc = [0, 0]  # 主块位置
    l_cur_loc = [0, 0]  # 左块位置
    r_cur_loc = [0, 0]  # 右块位置
    pre_rows = int(float(all_num) / float(b_bk_w_1 + b_bk_w_2 + b_bk_w_3))
    # 小报告厅预先遍历一次 TODO
    if mod == 'Small':  # 小报告厅
        l_cur_loc = [1, 0]
        if bool_spaced:
            left_first_row_num = 1
            l_start_loc = [0, 1]
        else:
            l_start_loc = [0, 2]
            left_first_row_num = 3
        for i in range(0, num_college):
            num = s_list[i][1]
            if num < left_first_row_num:
                continue
            bool_find = False
            num_s = num - left_first_row_num
            if (num_s + l_cur_loc[1]) % b_bk_w_1 == 0:
                bool_find = SpecialSeat('l', num_s, l_cur_loc)
                if bool_find:
                    s_list[i][2] = True
                    seating_result.append([s_list[i][0], 'l', l_start_loc, [l_cur_loc[0] - 1, b_bk_w_1 - 1]])
                    break
            l_start_loc2, bool_l = PreSeating('l', num_s, l_cur_loc)
            l_pre_loc = copy.deepcopy(l_start_loc2)
            bool_find2 = False
            for j in range(num_college - 1, i, -1):
                num2 = s_list[j][1]
                if (num2 + l_pre_loc[1]) % b_bk_w_1 == 0:
                    bool_find2 = SpecialSeat('l', num_s, l_pre_loc)
                    if bool_find2:  # 如能整排，返回连同第i个学院和第j个学院一起排的块号和位置
                        s_list[i][2] = True
                        s_list[j][2] = True
                        seating_result.append([s_list[i][0], 'l', l_start_loc,
                                               [l_start_loc2[0], l_start_loc2[1] - 1]])
                        seating_result.append([s_list[j][0], 'l', l_start_loc2, [l_pre_loc[0] - 1, b_bk_w_1 - 1]])
                        l_cur_loc = l_pre_loc
                        break
            if bool_find2:
                break

    # print('pre_rows: ', pre_rows)
    # print('all_num: ', all_num)
    # pre_rows -= 1
    for i in range(0, num_college):  # 也许用递归更好
        if s_list[i][2]:  # 若已使用，则跳过
            continue
        # 判断在哪一个板块排座
        s_list[i][2] = True
        l_start_loc = copy.deepcopy(l_cur_loc)
        m_start_loc = copy.deepcopy(m_cur_loc)
        r_start_loc = copy.deepcopy(r_cur_loc)
        block_num, bool_find = JudgeBlock(m_cur_loc, l_cur_loc, r_cur_loc, s_list[i][1])
        if bool_find:  # 如能整排，返回已排的块号和位置
            if block_num == 'l':
                seating_result.append([s_list[i][0], 'l', l_start_loc, [l_cur_loc[0] - 1, b_bk_w_1 - 1]])
            elif block_num == 'm':
                seating_result.append([s_list[i][0], 'm', m_start_loc, [m_cur_loc[0] - 1, b_bk_w_2 - 1]])
            elif block_num == 'r':
                seating_result.append([s_list[i][0], 'r', r_start_loc, [r_cur_loc[0] - 1, b_bk_w_3 - 1]])
            continue

        # 进行第二深度遍历，预固定上一深度中每个板块的排座
        l_start_loc2, bool_l = PreSeating('l', s_list[i][1], l_cur_loc)
        l_pre_loc = copy.deepcopy(l_start_loc2)
        m_start_loc2, bool_m = PreSeating('m', s_list[i][1], m_cur_loc)
        m_pre_loc = copy.deepcopy(m_start_loc2)
        r_start_loc2, bool_r = PreSeating('r', s_list[i][1], r_cur_loc)
        r_pre_loc = copy.deepcopy(r_start_loc2)
        bool_find2 = False
        for j in range(num_college - 1, i, -1):
            if s_list[j][2]:  # 若已使用，则跳过
                continue
            block_num, bool_find2 = JudgeBlock(m_pre_loc, l_pre_loc, r_pre_loc, s_list[j][1])
            if bool_find2:  # 如能整排，返回连同第i个学院和第j个学院一起排的块号和位置
                if (block_num == 'l') & bool_l:
                    s_list[j][2] = True
                    seating_result.append([s_list[i][0], 'l', l_start_loc,
                                           [l_start_loc2[0], l_start_loc2[1] - 1]])
                    seating_result.append([s_list[j][0], 'l', l_start_loc2, [l_pre_loc[0] - 1, b_bk_w_1 - 1]])
                    l_cur_loc = l_pre_loc
                    break
                elif (block_num == 'm') & bool_m:
                    s_list[j][2] = True
                    seating_result.append([s_list[i][0], 'm', m_start_loc,
                                           [m_start_loc2[0], m_start_loc2[1] - 1]])
                    seating_result.append([s_list[j][0], 'm', m_start_loc2, [m_pre_loc[0] - 1, b_bk_w_2 - 1]])
                    m_cur_loc = m_pre_loc
                    break
                elif (block_num == 'r') & bool_r:
                    s_list[j][2] = True
                    seating_result.append([s_list[i][0], 'r', r_start_loc,
                                           [r_start_loc2[0], r_start_loc2[1] - 1]])
                    seating_result.append([s_list[j][0], 'r', r_start_loc2, [r_pre_loc[0] - 1, b_bk_w_3 - 1]])
                    r_cur_loc = r_pre_loc
                    break
        if bool_find2:
            continue
        if bool_m:
            m_cur_loc = m_start_loc2
            seating_result.append([s_list[i][0], 'm', m_start_loc, [m_start_loc2[0], m_start_loc2[1] - 1]])
        elif bool_l:
            l_cur_loc = l_start_loc2
            seating_result.append([s_list[i][0], 'l', l_start_loc, [l_start_loc2[0], l_start_loc2[1] - 1]])
        elif bool_r:
            r_cur_loc = r_start_loc2
            seating_result.append([s_list[i][0], 'r', r_start_loc, [r_start_loc2[0], r_start_loc2[1] - 1]])
        else:
            add_rows = int((s_list[i][1] + m_cur_loc[1]) / b_bk_w_2)
            remand_cols = (s_list[i][1] + m_cur_loc[1]) % b_bk_w_2
            m_cur_loc[0] = m_cur_loc[0] + add_rows
            m_cur_loc[1] = remand_cols
            seating_result.append([s_list[i][0], 'm', m_start_loc, [m_cur_loc[0], m_cur_loc[1] - 1]])
    # for i in range(0, num_college):
    #     print(seating_result[i])
    return seating_result


def PreSeating(sign, num, cur_loc):
    bool_b = False
    start_loc = copy.deepcopy(cur_loc)
    if sign == 'l':
        b_bk_w = b_bk_w_1
        upper_rows = pre_rows - det_rows + 1
    elif sign == 'm':
        b_bk_w = b_bk_w_2
        upper_rows = pre_rows + det_rows + 1
    else:
        upper_rows = pre_rows - det_rows + 1
        b_bk_w = b_bk_w_3
    add_rows = int((num + start_loc[1]) / b_bk_w)
    remand_cols = (num + start_loc[1]) % b_bk_w
    if (start_loc[0] + add_rows) <= upper_rows:
        start_loc[0] = start_loc[0] + add_rows
        start_loc[1] = remand_cols
        bool_b = True
    return start_loc, bool_b


def SpecialSeat(sign, num, cur_loc):
    bool_find = False
    if sign == 'l':
        b_bk_w = b_bk_w_1
        upper_rows = pre_rows - det_rows + 1
    elif sign == 'm':
        b_bk_w = b_bk_w_2
        upper_rows = pre_rows + det_rows + 1
    else:
        upper_rows = pre_rows - det_rows + 1
        b_bk_w = b_bk_w_3
    add_rows = int((num + cur_loc[1]) / b_bk_w)
    if (cur_loc[0] + add_rows) <= upper_rows:
        cur_loc[0] = cur_loc[0] + add_rows
        cur_loc[1] = 0
        bool_find = True
    return bool_find


def JudgeBlock(m_cur_loc, l_cur_loc, r_cur_loc, num):
    bool_find = False
    block_num = 'm'
    if m_cur_loc[0] <= (pre_rows + det_rows):
        if (num + l_cur_loc[1]) % b_bk_w_1 == 0:
            bool_find = SpecialSeat('l', num, l_cur_loc)
            block_num = 'l'
        if (not bool_find) & ((num + m_cur_loc[1]) % b_bk_w_2 == 0):
            bool_find = SpecialSeat('m', num, m_cur_loc)
            block_num = 'm'
        if (not bool_find) & ((num + r_cur_loc[1]) % b_bk_w_3 == 0):
            bool_find = SpecialSeat('r', num, r_cur_loc)
            block_num = 'r'
    elif l_cur_loc[0] <= (pre_rows - det_rows):
        block_num = 'l'
        if (num + l_cur_loc[1]) % b_bk_w_1 == 0:
            bool_find = SpecialSeat('l', num, l_cur_loc)
            block_num = 'l'
        if (not bool_find) & ((num + r_cur_loc[1]) % b_bk_w_3 == 0):
            bool_find = SpecialSeat('r', num, r_cur_loc)
            block_num = 'r'
    elif r_cur_loc[0] <= (pre_rows - det_rows):
        if (num + r_cur_loc[1]) % b_bk_w_3 == 0:
            bool_find = SpecialSeat('r', num, r_cur_loc)
        block_num = 'r'

    return block_num, bool_find


def SortList(final_list):
    global all_num
    all_num = 0
    s_list = []
    while len(final_list) != 0:
        max_idx = 0
        max_num = 0
        for i in range(0, len(final_list)):
            if final_list[i][1] > max_num:
                max_idx = i
                max_num = final_list[i][1]
        s_list.append(final_list[max_idx])
        s_list[-1].append(False)
        all_num += final_list[max_idx][1]
        del final_list[max_idx]
    return s_list


if __name__ == '__main__':
    f_list = [['能源学院', 3], ['公共卫生学院', 6], ['航空航天学院', 15], ['海外教育学院', 2], ['国际学院', 2], ['化学化工学院', 4], ['药学院', 4],
              ['萨本栋微米纳米科学技术研究院', 2], ['电子科学与技术学院', 10], ['生命科学学院', 11], ['医学院', 10], ['信息学院', 15], ['环境与生态学院', 6],
              ['海洋与地球学院', 8], ['少数民族预科班', 2]]
    InBigHall(' ', f_list, 2, ' ', 1)
    # print(['l'*23])
    # zero_list = [[0 for val in range(5)] for val2 in range(21)]
    # print(zero_list)
    # a = 13
    # b = 2
    # print((a - a % b) / b)
    # print(int(a) / int(b))
    # print(a % b)
