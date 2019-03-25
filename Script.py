 # *- coding: utf-8 -*-
"""
Spyder Editor

This script file created by Vanna
2019.3.2
"""
import pandas as pd;
import numpy as np

'''
学生数据导入
'''
Student_info = pd.read_excel('D:\Mass\Sports_data\studentlist\CLASS034.xlsx')
#删除空值列
#Student_info = Student_info.drop(columns=['Prog.','Major','Spec','Cur. status','Remarks'])

# #学生信息清洗
#打印出列名
print(Student_info.columns.values);
#列名更改
Student_info.rename(columns={ Student_info.columns[0]: "Student_ID" }, inplace=True)
Student_info.rename(columns={ Student_info.columns[1]: "Eng_Name" }, inplace=True);
Student_info.rename(columns={ Student_info.columns[2]: "Ch_Name" }, inplace=True)
#Student_ID转换成全大写
Student_info["Student_ID"]=Student_info["Student_ID"].str.upper()
Student_info["Student_ID"]=Student_info["Student_ID"].str.replace("-","")
#字符串的空格改写
Student_info["Ch_Name"]=Student_info["Ch_Name"].str.replace(" ","")
#学号重整
Student_info["Student_ID"]=Student_info["Student_ID"].str.replace("AB","B")
Student_info["Student_ID"]=Student_info["Student_ID"].str.replace("BB","B")
Student_info["Student_ID"]=Student_info["Student_ID"].str.replace("DB","B")
Student_info["Student_ID"]=Student_info["Student_ID"].str.replace("HB","B")
Student_info["Student_ID"]=Student_info["Student_ID"].str.replace("SB","B")

'''
活动数据导入

第二种数据格式分解-导入
Active_pre_in = pd.read_excel('D:\Mass\Sports_data\Active\A_In.xlsx')
Active_pre_out = pd.read_excel('D:\Mass\Sports_data\Active\A_Out.xlsx')
Active_pre_data = Active_pre_in.join(Active_pre_out.set_index('Student_ID'),on='Student_ID')

'''

'''
A_JunKR201901
'''
A_JunKR = pd.read_excel('D:\Mass\Sports_data\Active\A_Keep Running Project -FED_201901January.xlsx')

# #活动信息清洗
#列名更改
A_JunKR.rename(columns={A_JunKR.columns[0]: "Student_ID" }, inplace=True)
A_JunKR.rename(columns={A_JunKR.columns[1]:"JunKR_In"}, inplace =True)
A_JunKR.rename(columns={A_JunKR.columns[2]:"JunKR_Out"}, inplace =True)

#时间差值计算
A_JunKR['January_KeepRun']=A_JunKR['JunKR_Out']-A_JunKR['JunKR_In']
#转换为float型的分钟
A_JunKR['January_KeepRun']=A_JunKR['January_KeepRun'] / np.timedelta64(60,'s')

'''
A_UMOpening20181024

A_UMOpening = pd.read_excel('D:\Mass\Sports_data\Active\A_UM Sports Day Opening.xlsx')
A_UMOpening.rename(columns={A_UMOpening.columns[0]: "Student_ID" }, inplace=True)
A_UMOpening.rename(columns={A_UMOpening.columns[1]:"UMOpening_In"}, inplace =True)
A_UMOpening.rename(columns={A_UMOpening.columns[2]:"UMOpening_Out"}, inplace =True)

A_UMOpening['UMOpening']=A_UMOpening['UMOpening_Out']-A_UMOpening['UMOpening_In']
A_UMOpening['UMOpening']=A_UMOpening['UMOpening'] / np.timedelta64(60,'s')
'''
'''
A_NightRun20190117
'''
A_NightRun0117 =pd.read_excel('D:\Mass\Sports_data\Active\A_NightRun20190117.xlsx')
A_NightRun0117.rename(columns={A_NightRun0117.columns[0]:"Student_ID"}, inplace=True)
A_NightRun0117.rename(columns={A_NightRun0117.columns[1]:"NightRun_In0117"}, inplace=True )
A_NightRun0117.rename(columns={A_NightRun0117.columns[2]:"NightRun_Out0117"},inplace=True)

A_NightRun0117['NightRun0117']=A_NightRun0117['NightRun_Out0117']-A_NightRun0117['NightRun_In0117']
A_NightRun0117['NightRun0117']=A_NightRun0117['NightRun0117'] / np.timedelta64(60,'s')



'''
A_NightRun20190228
'''
A_NightRun0228 =pd.read_excel('D:\Mass\Sports_data\Active\A_NightRun20190228.xlsx')
A_NightRun0228.rename(columns={A_NightRun0228.columns[0]:"Student_ID"}, inplace=True)
A_NightRun0228.rename(columns={A_NightRun0228.columns[1]:"NightRun_In0228"}, inplace=True )
A_NightRun0228.rename(columns={A_NightRun0228.columns[2]:"NightRun_Out0228"},inplace=True)

A_NightRun0228['NightRun0228']=A_NightRun0228['NightRun_Out0228']-A_NightRun0228['NightRun_In0228']
A_NightRun0228['NightRun0228']=A_NightRun0228['NightRun0228'] / np.timedelta64(60,'s')

'''
    A_BasketGame0309
'''
A_BasketGame0309 =pd.read_excel('D:\Mass\Sports_data\Active\A_BasketGame0309.xlsx')
A_BasketGame0309.rename(columns={A_BasketGame0309.columns[0]:"Student_ID"}, inplace=True)
A_BasketGame0309.rename(columns={A_BasketGame0309.columns[1]:"BasketGame_In0309"}, inplace=True )
A_BasketGame0309.rename(columns={A_BasketGame0309.columns[2]:"BasketGame_Out0309"},inplace=True)

A_BasketGame0309['BasketGame0309']=A_BasketGame0309['BasketGame_Out0309']-A_BasketGame0309['BasketGame_In0309']
A_BasketGame0309['BasketGame0309']=A_BasketGame0309['BasketGame0309'] / np.timedelta64(60,'s')

'''
    A_BasketGame0310
'''
A_BasketGame0310 =pd.read_excel('D:\Mass\Sports_data\Active\A_BasketGame0310.xlsx')
A_BasketGame0310.rename(columns={A_BasketGame0310.columns[0]:"Student_ID"}, inplace=True)
A_BasketGame0310.rename(columns={A_BasketGame0310.columns[1]:"BasketGame_In0310"}, inplace=True )
A_BasketGame0310.rename(columns={A_BasketGame0310.columns[2]:"BasketGame_Out0310"},inplace=True)

A_BasketGame0310['BasketGame0310']=A_BasketGame0310['BasketGame_Out0310']-A_BasketGame0310['BasketGame_In0310']
A_BasketGame0310['BasketGame0310']=A_BasketGame0310['BasketGame0310'] / np.timedelta64(60,'s')

'''
A_Swim0219
'''
A_Swim0219 =pd.read_excel('D:\Mass\Sports_data\Active\A_Swim0219.xlsx')
A_Swim0219.rename(columns={A_Swim0219.columns[0]:"Student_ID"}, inplace=True)
A_Swim0219.rename(columns={A_Swim0219.columns[1]:"Swim_In0219"}, inplace=True )
A_Swim0219.rename(columns={A_Swim0219.columns[2]:"Swim_Out0219"},inplace=True)

A_Swim0219['Swim0219']=A_Swim0219['Swim_Out0219']-A_Swim0219['Swim_In0219']
A_Swim0219['Swim0219']=A_Swim0219['Swim0219'] / np.timedelta64(60,'s')



'''

'''
#A_DiaboloSports1024
'''
A_Diabolo1024 =pd.read_excel('D:\Mass\Sports_data\Active\A_DiaboloSports1024.xlsx')
A_Diabolo1024.rename(columns={A_Diabolo1024.columns[0]:"Student_ID"}, inplace=True)
A_Diabolo1024.rename(columns={A_Diabolo1024.columns[1]:"Diabolo_In"}, inplace=True )
A_Diabolo1024.rename(columns={A_Diabolo1024.columns[2]:"Diabolo_Out"},inplace=True)

A_Diabolo1024['Diabolo1024']=A_Diabolo1024['Diabolo_Out']-A_Diabolo1024['Diabolo_In']
A_Diabolo1024['Diabolo1024']=A_Diabolo1024['Diabolo1024'] / np.timedelta64(60,'s')


'''
#A_DiaboloSports1027
'''
A_Diabolo1027 =pd.read_excel('D:\Mass\Sports_data\Active\A_DiaboloSport1027.xlsx')
A_Diabolo1027.rename(columns={A_Diabolo1027.columns[0]:"Student_ID"}, inplace=True)
A_Diabolo1027.rename(columns={A_Diabolo1027.columns[1]:"Diabolo_In1027"}, inplace=True )
A_Diabolo1027.rename(columns={A_Diabolo1027.columns[2]:"Diabolo_Out1027"},inplace=True)

A_Diabolo1027['Diabolo1027']=A_Diabolo1027['Diabolo_Out1027']-A_Diabolo1027['Diabolo_In1027']
A_Diabolo1027['Diabolo1027']=A_Diabolo1027['Diabolo1027'] / np.timedelta64(60,'s')

'''
#A_DiaboloSports1107
'''
A_Diabolo1107 =pd.read_excel('D:\Mass\Sports_data\Active\A_DiaboloSport1107.xlsx')
A_Diabolo1107.rename(columns={A_Diabolo1107.columns[0]:"Student_ID"}, inplace=True)
A_Diabolo1107.rename(columns={A_Diabolo1107.columns[1]:"Diabolo_In1107"}, inplace=True )
A_Diabolo1107.rename(columns={A_Diabolo1107.columns[2]:"Diabolo_Out1107"},inplace=True)

A_Diabolo1107['Diabolo1107']=A_Diabolo1107['Diabolo_Out1107']-A_Diabolo1107['Diabolo_In1107']
A_Diabolo1107['Diabolo1107']=A_Diabolo1107['Diabolo1107'] / np.timedelta64(60,'s')

'''
#A_DiaboloSports1114
'''
A_Diabolo1114 =pd.read_excel('D:\Mass\Sports_data\Active\A_DiaboloSport1114.xlsx')
A_Diabolo1114.rename(columns={A_Diabolo1114.columns[0]:"Student_ID"}, inplace=True)
A_Diabolo1114.rename(columns={A_Diabolo1114.columns[1]:"Diabolo_In1114"}, inplace=True )
A_Diabolo1114.rename(columns={A_Diabolo1114.columns[2]:"Diabolo_Out1114"},inplace=True)

A_Diabolo1114['Diabolo1114']=A_Diabolo1114['Diabolo_Out1114']-A_Diabolo1114['Diabolo_In1114']
A_Diabolo1114['Diabolo1114']=A_Diabolo1114['Diabolo1114'] / np.timedelta64(60,'s')

'''
#A_DiaboloSports1121
'''
A_Diabolo1121 =pd.read_excel('D:\Mass\Sports_data\Active\A_DiaboloSport1121.xlsx')
A_Diabolo1121.rename(columns={A_Diabolo1121.columns[0]:"Student_ID"}, inplace=True)
A_Diabolo1121.rename(columns={A_Diabolo1121.columns[1]:"Diabolo_In1121"}, inplace=True )
A_Diabolo1121.rename(columns={A_Diabolo1121.columns[2]:"Diabolo_Out1121"},inplace=True)

A_Diabolo1121['Diabolo1121']=A_Diabolo1121['Diabolo_Out1121']-A_Diabolo1121['Diabolo_In1121']
A_Diabolo1121['Diabolo1121']=A_Diabolo1121['Diabolo1121'] / np.timedelta64(60,'s')

'''
'''
数据整合
'''
#表的关联
Reletive = Student_info.join(A_JunKR.set_index('Student_ID'),on='Student_ID',how='left')
#Reletive = Reletive.join(A_UMOpening.set_index('Student_ID'),on='Student_ID',how='left')
Reletive = Reletive.join(A_NightRun0117.set_index('Student_ID'),on='Student_ID',how='left')
Reletive = Reletive.join(A_NightRun0228.set_index('Student_ID'),on='Student_ID',how='left')
Reletive = Reletive.join(A_BasketGame0309.set_index('Student_ID'),on='Student_ID',how='left')
Reletive = Reletive.join(A_Swim0219.set_index('Student_ID'),on='Student_ID',how='left')
Reletive = Reletive.join(A_BasketGame0310.set_index('Student_ID'),on='Student_ID',how='left')

#Reletive = Reletive.join(A_Diabolo1024.set_index('Student_ID'),on='Student_ID',how='left')
#Reletive = Reletive.join(A_Diabolo1027.set_index('Student_ID'),on='Student_ID',how='left')
#Reletive = Reletive.join(A_Diabolo1107.set_index('Student_ID'),on='Student_ID',how='left')
#Reletive = Reletive.join(A_Diabolo1114.set_index('Student_ID'),on='Student_ID',how='left')
#Reletive = Reletive.join(A_Diabolo1121.set_index('Student_ID'),on='Student_ID',how='left')


'''
暂储数据
'''
Reletive.to_excel("D:\Mass\Sports_data\Combine2019S1\CPED034.xlsx")
