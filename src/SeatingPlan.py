import tkinter
from tkinter import *
from LoadAndSave import Get_Setting, Save_Setting, SaveToYaml, LoadFromYaml
import os
from Seating import InBigHall, InSmallHall
from LoadAndSave import CreateNewFile

setting_list = []
init_path = os.getcwd()
setting_init_path = os.path.join(init_path, 'Setting')


class SP_GUI:
    # all_entry = []
    final_dis_list = []
    bools = []

    def __init__(self):
        self.sp_win = Tk()
        self.sp_win.title('这是一个内部使用的简单排座系统')
        self.sp_win.geometry('1150x700')  # 窗口大小

        # 配置信息显示
        label1 = Label(self.sp_win, text='分配前人数', font=('宋体', 10), width=30, height=1)
        label1.grid(row=0, column=1, padx=5, pady=5)
        label2 = Label(self.sp_win, text='分配后人数', font=('宋体', 10), width=30, height=1)
        label2.grid(row=0, column=2, padx=5, pady=5)
        self.college_size = len(setting_list)
        self.before_dis_entry = []
        self.after_dis_entry = []
        all_num = 0
        for i in range(0, self.college_size):
            self.bools.append(IntVar())
            self.bools[-1].set(1)
            college_c = Checkbutton(self.sp_win, text=setting_list[i][0], variable=self.bools[-1], onvalue=1,
                                    offvalue=0, font=('宋体', 10), width=30, height=1)
            college_c.grid(row=i + 1, column=0, padx=5, pady=5)
            # college_l = Label(self.sp_win, text=setting_list[i][0], font=('宋体', 10), width=30, height=1)
            # college_l.grid(row=i+1, column=0, padx=5, pady=5)
            number_e = Entry(self.sp_win)
            number_e.grid(row=i + 1, column=1, padx=5, pady=5)
            number_e.insert(0, str(setting_list[i][1]))
            self.before_dis_entry.append(number_e)
            number_e2 = Entry(self.sp_win)
            number_e2.grid(row=i + 1, column=2, padx=5, pady=5)
            number_e2.insert(0, str(setting_list[i][1]))
            self.after_dis_entry.append(number_e2)
            all_num = all_num + setting_list[i][1]
        all_college_l = Label(self.sp_win, text='总人数', font=('宋体', 15), width=25, height=1)
        all_college_l.grid(row=self.college_size + 1, column=0, padx=5, pady=5)
        all_number_e = Entry(self.sp_win)
        all_number_e.grid(row=self.college_size + 1, column=1, padx=5, pady=5)
        all_number_e.insert(0, str(all_num))
        self.before_dis_entry.append(all_number_e)
        all_number_e2 = Entry(self.sp_win)
        all_number_e2.grid(row=self.college_size + 1, column=2, padx=5, pady=5)
        all_number_e2.insert(0, str(all_num))
        self.after_dis_entry.append(all_number_e2)

        # 按键
        dis_button = Button(self.sp_win, text="分配人数", command=self.Distribute_Number, width=15)
        default_setting = Button(self.sp_win, text="默认人数", command=self.Load_Default, width=15)
        calculate = Button(self.sp_win, text="计算总人数", command=self.Calculate_All_Number, width=15)
        choose_scene1 = Button(self.sp_win, text="大报告厅排座", command=self.SeatingInBigHall, width=15)
        choose_scene2 = Button(self.sp_win, text="小报告厅排座", command=self.SeatingInSmallHall, width=15)
        open_result = Button(self.sp_win, text="打开结果", command=OpenResult, width=15)
        close_win = Button(self.sp_win, text="关闭", font=('宋体', 20), command=self.CloseAndSave, width=45)
        # 控件显示
        dis_button.grid(row=self.college_size + 2, column=2, padx=5, pady=5)
        default_setting.grid(row=self.college_size + 2, column=1, padx=5, pady=5)
        calculate.grid(row=self.college_size + 2, column=0, padx=5, pady=5)
        close_win.grid(row=self.college_size + 3, column=1, columnspan=3, padx=5, pady=5)
        choose_scene1.grid(row=self.college_size + 2, column=3, padx=5, pady=5)
        choose_scene2.grid(row=self.college_size + 1, column=3, padx=5, pady=5)
        open_result.grid(row=self.college_size + 2, column=4, padx=5, pady=5)

        # 获取输入信息
        info = LoadFromYaml()
        # print(info)
        dis_l = Label(self.sp_win, text='分配总人数', font=('宋体', 12), width=25, height=1)
        dis_l.grid(row=1, column=3, padx=5, pady=5)
        num_to_dis_e = Entry(self.sp_win)
        num_to_dis_e.grid(row=1, column=4, padx=10, pady=10)
        num_to_dis_e.insert(0, info['分配总人数'])
        self.num_to_dis_e = num_to_dis_e
        min_num_l = Label(self.sp_win, text='学院最少人数', font=('宋体', 12), width=25, height=1)
        min_num_l.grid(row=2, column=3, padx=5, pady=5)
        min_num_e = Entry(self.sp_win)
        min_num_e.grid(row=2, column=4, padx=10, pady=10)
        min_num_e.insert(0, info['学院最少人数'])
        self.min_num_e = min_num_e
        name_l = Label(self.sp_win, text='活动名称', font=('宋体', 12), width=25, height=1)
        name_l.grid(row=3, column=3, padx=5, pady=5)
        name_e = Entry(self.sp_win)
        name_e.grid(row=3, column=4, padx=10, pady=10)
        name_e.insert(0, info['活动名称'])
        self.name_e = name_e
        date_l = Label(self.sp_win, text='活动时间', font=('宋体', 12), width=25, height=1)
        date_l.grid(row=4, column=3, padx=5, pady=5)
        date_e = Entry(self.sp_win)
        date_e.grid(row=4, column=4, padx=10, pady=10)
        date_e.insert(0, info['活动时间'])
        self.date_e = date_e
        vip_l = Label(self.sp_win, text='嘉宾排数', font=('宋体', 12), width=25, height=1)
        vip_l.grid(row=5, column=3, padx=5, pady=5)
        vip_e = Entry(self.sp_win)
        vip_e.grid(row=5, column=4, padx=10, pady=10)
        vip_e.insert(0, info['嘉宾排数'])
        self.vip_e = vip_e

        self.spaced_v = IntVar()
        self.spaced_v.set(info['隔座'])
        spaced_r = Checkbutton(self.sp_win, text='隔座', variable=self.spaced_v, onvalue=1,
                               offvalue=0, font=('宋体', 12), width=25, height=1)
        spaced_r.grid(row=6, column=4, padx=5, pady=5)

        self.sp_win.mainloop()

    def Calculate_All_Number(self):
        all_num = 0
        all_dis_num = 0
        for i in range(0, self.college_size):
            if self.bools[i].get() == 1:
                all_num = all_num + int(self.before_dis_entry[i].get())
                all_dis_num = all_dis_num + int(self.after_dis_entry[i].get())

        self.before_dis_entry[-1].delete(0, 'end')
        self.before_dis_entry[-1].insert(0, str(all_num))
        self.after_dis_entry[-1].delete(0, 'end')
        self.after_dis_entry[-1].insert(0, str(all_dis_num))

    def Load_Default(self):
        global setting_list
        setting_list = Get_Setting(setting_init_path, 'DEFAULT')
        all_num = 0
        for i in range(0, self.college_size):
            self.before_dis_entry[i].delete(0, 'end')
            self.before_dis_entry[i].insert(0, str(setting_list[i][1]))
            all_num = all_num + setting_list[i][1]
        self.before_dis_entry[-1].delete(0, 'end')
        self.before_dis_entry[-1].insert(0, str(all_num))

    def Distribute_Number(self):
        # self.final_dis_list.clear()
        global setting_list
        all_num = 0  # 当前界面中各学院的总人数
        every_num = []  # 当前界面中各学院人数
        all_num_to_dis = float(self.num_to_dis_e.get())  # 预分配的总人数
        min_num = self.min_num_e.get()  # 最小分配到的人数
        every_num_after_dis = []  # 勾选各学院的分配人数
        every_idx_after_dis = []  # 勾选的学院索引
        all_num_after_dis = 0  # 分配后的总人数，最终和预分配人数相等
        for i in range(0, self.college_size):
            str_num_value = self.before_dis_entry[i].get()
            setting_list[i][1] = int(str_num_value)
            if self.bools[i].get() == 1:
                all_num = all_num + float(str_num_value)
                every_num.append(float(str_num_value))

                every_idx_after_dis.append([i])
        j = 0
        for i in range(0, self.college_size):
            if self.bools[i].get() == 1:
                after_dis_num = round(every_num[j] * all_num_to_dis / all_num)
                j = j + 1
                if int(after_dis_num) < int(min_num):
                    every_num_after_dis.append(int(min_num))
                else:
                    every_num_after_dis.append(int(after_dis_num))
                all_num_after_dis = all_num_after_dis + every_num_after_dis[-1]

        # all_num_after_dis = int(all_num_after_dis)
        while all_num_after_dis != int(all_num_to_dis):
            if all_num_after_dis > int(all_num_to_dis):
                max_num = max(every_num_after_dis)
                max_id = every_num_after_dis.index(max_num)
                every_num_after_dis[max_id] = every_num_after_dis[max_id] - 1
                all_num_after_dis = all_num_after_dis - 1
            else:
                min_num = min(every_num_after_dis)
                min_id = every_num_after_dis.index(min_num)
                every_num_after_dis[min_id] = every_num_after_dis[min_id] + 1
                all_num_after_dis = all_num_after_dis + 1

        j = 0
        for i in range(0, self.college_size):
            self.after_dis_entry[i].delete(0, 'end')
            if self.bools[i].get() == 1:
                self.after_dis_entry[i].insert(0, str(every_num_after_dis[j]))
                # self.final_dis_list.append([setting_list[i][0], every_num_after_dis[j]])
                j = j + 1
            else:
                self.after_dis_entry[i].insert(0, '0')
                # self.final_dis_list.append([setting_list[i][0], 0])
        # for i in range(0, len(self.final_dis_list)):
        #     print(self.final_dis_list[i])
        # print(self.final_dis_list)

        self.before_dis_entry[-1].delete(0, 'end')
        self.before_dis_entry[-1].insert(0, str(int(all_num)))
        self.after_dis_entry[-1].delete(0, 'end')
        self.after_dis_entry[-1].insert(0, str(all_num_after_dis))

    def CloseAndSave(self):
        save_to_yaml = {'分配总人数': self.num_to_dis_e.get(),
                        '学院最少人数': self.min_num_e.get(),
                        '活动名称': self.name_e.get(),
                        '活动时间': self.date_e.get(),
                        '嘉宾排数': self.vip_e.get(),
                        '隔座': self.spaced_v.get()}
        # print(self.spaced_v.get())
        SaveToYaml(save_to_yaml)
        Save_Setting(setting_init_path, setting_list)
        self.sp_win.destroy()

    def GetNowDis(self):
        self.final_dis_list.clear()
        for i in range(0, self.college_size):
            if self.bools[i].get() == 0:
                continue
            val = int(self.after_dis_entry[i].get())
            if val == 0:
                continue
            self.final_dis_list.append([setting_list[i][0], val])

    def SeatingInBigHall(self):
        vip_rows = int(self.vip_e.get())
        empty_excel_path = os.path.join(setting_init_path, '翔安图书馆大报告厅座位图.xlsx')
        # print(empty_excel_path)
        result_path = os.path.join(init_path, 'Result',
                                   self.date_e.get() + self.name_e.get() + '.xlsx')
        # print(result_path)
        # CreateNewFile(empty_excel_path, result_path)
        self.GetNowDis()
        final_list = self.final_dis_list
        spaced = self.spaced_v.get()
        sign = InBigHall(result_path, final_list, vip_rows, self.name_e.get(), spaced)
        if not sign:
            label_false = Label(self.sp_win, text='人数过多，请调整！！！', font=('宋体', 20), width=30, height=1)
            label_false.grid(row=7, column=3, rowspan=3, columnspan=2, padx=5, pady=5)
        else:
            label_false = Label(self.sp_win, text='排座完成！', font=('宋体', 20), width=30, height=1)
            label_false.grid(row=7, column=3, rowspan=3, columnspan=2, padx=5, pady=5)

    def SeatingInSmallHall(self):
        vip_rows = int(self.vip_e.get())
        empty_excel_path = os.path.join(setting_init_path, '翔安图书馆小报告厅座位图.xlsx')
        result_path = os.path.join(init_path, 'Result',
                                   self.date_e.get() + self.name_e.get() + '.xlsx')
        # CreateNewFile(empty_excel_path, result_path)
        self.GetNowDis()
        final_list = self.final_dis_list
        spaced = self.spaced_v.get()
        sign = InSmallHall(result_path, final_list, vip_rows, self.name_e.get(), spaced)
        if not sign:
            label_false = Label(self.sp_win, text='人数过多，请调整！！！', font=('宋体', 20), width=30, height=1)
            label_false.grid(row=7, column=3, rowspan=3, columnspan=2, padx=5, pady=5)
        else:
            label_false = Label(self.sp_win, text='排座完成！', font=('宋体', 20), width=30, height=1)
            label_false.grid(row=7, column=3, rowspan=3, columnspan=2, padx=5, pady=5)


def OpenResult():
    path = os.path.join(init_path, 'Result')
    os.startfile(path)


if __name__ == '__main__':
    setting_list = Get_Setting(setting_init_path, 'LAST')
    SP_GUI()

    # window = Tk()  # 一个窗口
    # window.title('my window')  # 标题
    # window.geometry('200x200')  # 长和宽
    #
    # var = StringVar()  # 一个字符变量
    # l = Label(window, bg='yellow', width=20, text='empty')  # l是一个标签
    # l.pack()
    #
    #
    # def print_selection():
    #     l.config(text='you have selected' + var.get())  # config是修改参数的函数
    #
    #
    # # 三个选择按钮
    # r1 = Radiobutton(window, text='Option A', variable=var, value='A',
    #                     command=print_selection)  # 当你选择该选项时, var就会被赋值为A
    # r1.pack()
    # r2 = Radiobutton(window, text='Option B', variable=var, value='B',
    #                     command=print_selection)  # 当你选择该选项时, var就会被赋值为B
    # r2.pack()
    # r3 = Radiobutton(window, text='Option C', variable=var, value='C',
    #                     command=print_selection)  # 当你选择该选项时, var就会被赋值为C
    # r3.pack()
    #
    # window.mainloop()
