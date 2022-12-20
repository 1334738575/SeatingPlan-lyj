# SeatingPlan-lyj
这是一个内部使用的简单排座脚本！
双击运行SeatingPlan.exe， 即可使用！
1.主界面中第一列为各个学院的学院名和总人数，勾选该学院表示在分配人数时，该学院按比例进行分配；
2.点击“计算总人数”，则会进行计算“分配前人数”和“分配后人数”的总和；
3.第二列为分配前的各学院人数，每次打开会从Setting文件的last.txt文档中读取上一次设置的人数，可以人为进行修改，分配人数时以该列为准；
3.点击“默认人数”，系统会在Setting文件夹的default.txt文档中读取默认人数文件；
4.第三列为分配后的人数，可以人为修改，排座时以该列为准；
5.点击“分配人数”，系统根据学院的勾选情况，按照“分配总人数”进行比例分配；
6.第五列为排座的配置参数，每次运行会自动加载上一次运行的结果（保存在Setting文件夹中的last_config.yaml文档中），其中：
	i.“分配总人数”为需要进行分配的人数（不包括某些学院已固定的人数）；
	ii.“学院最少人数”即为分配人数时，每个学院的最少人数（不包括未勾选的学院）；
	iii.“活动名称”和“活动时间”不能同时为空；
	iv.“嘉宾排数”即为嘉宾所占排数，不能为小数；
	v.勾选“隔座”即可进行隔座排座。
7.点击“小报告厅排座”和“大报告厅排座”即可进行对应模板的排座（当总人数超过报告厅所能容纳的最大人数时，会提示“人数过多，请调整！！！”，而排座成功时，会提示“排座完成！”）；
8.点击“打开结果”，系统会打开exe目录下的Result文件夹，“.xlsx”保存在里面；
9.点击“关闭”为正常退出，此时系统会保存当前各学院人数至Setting文件夹的last.txt文档中，保存当前配置参数至Setting文件夹中的last_config.yaml文档中。
10.排座效果出现小问题时，对“.xlsx”进行简单修改即可，如果出现严重bug，或者排座效果极差，亦或有其他建设性意见，请发邮件至：1334738575@qq.com，十分感谢！