# -*- coding: utf-8 -*-
"""
Created on Tue Apr  9 23:14:50 2019

@author: 马文越
"""

 # *- coding: utf-8 -*-
"""
Spyder Editor

This script file created by Vanna
2019.3.2
"""
import pandas as pd;
import numpy as np
import os

'''
学生数据导入
'''
#Student_info = pd.read_excel('D:\Mass\Sports_data\studentlist\CLASS022.xlsx')
Class_dir='D:\Mass\Sports_data\studentlist'
Combine_dir='D:\Mass\Sports_data\Combine2019S1\stu_final'
for file in os.listdir(Class_dir) :
   if file.endswith('.xlsx'):
        file_name = Class_dir + '/' + file
      
        Student_info = pd.read_excel(file_name)
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
        Student_info["Student_ID"]=Student_info["Student_ID"].str.replace("CB","B")
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
        A_BasketGame0309
        '''
        A_BasketGame0309 =pd.read_excel('D:\Mass\Sports_data\Active\Comp\A_BasketGame0309.xlsx')
        A_BasketGame0309 =A_BasketGame0309.drop_duplicates(subset = None, keep = 'last')
        A_BasketGame0309.rename(columns={A_BasketGame0309.columns[0]:"Student_ID"}, inplace=True)
        A_BasketGame0309.rename(columns={A_BasketGame0309.columns[1]:"BasketGame_In0309"}, inplace=True )
        A_BasketGame0309.rename(columns={A_BasketGame0309.columns[2]:"BasketGame_Out0309"},inplace=True)
        
        A_BasketGame0309['BasketGame0309']=A_BasketGame0309['BasketGame_Out0309']-A_BasketGame0309['BasketGame_In0309']
        A_BasketGame0309['BasketGame0309']=A_BasketGame0309['BasketGame0309'] / np.timedelta64(60,'s')
        
        '''
        A_BasketGame0310
        '''
        A_BasketGame0310 =pd.read_excel('D:\Mass\Sports_data\Active\Comp\A_BasketGame0310.xlsx')
        A_BasketGame0310 =A_BasketGame0310.drop_duplicates(subset = None, keep = 'last')
        A_BasketGame0310.rename(columns={A_BasketGame0310.columns[0]:"Student_ID"}, inplace=True)
        A_BasketGame0310.rename(columns={A_BasketGame0310.columns[1]:"BasketGame_In0310"}, inplace=True )
        A_BasketGame0310.rename(columns={A_BasketGame0310.columns[2]:"BasketGame_Out0310"},inplace=True)
        
        A_BasketGame0310['BasketGame0310']=A_BasketGame0310['BasketGame_Out0310']-A_BasketGame0310['BasketGame_In0310']
        A_BasketGame0310['BasketGame0310']=A_BasketGame0310['BasketGame0310'] / np.timedelta64(60,'s')
        
        '''
        A_BasketGame0316
        '''
        A_BasketGame0316 = pd.read_excel('D:\Mass\Sports_data\Active\Comp\A_BasketGame0316.xlsx')
        A_BasketGame0316 =A_BasketGame0316.drop_duplicates(subset = None, keep = 'last')
        A_BasketGame0316.rename(columns={A_BasketGame0316.columns[0]: "Student_ID" }, inplace=True)
        A_BasketGame0316.rename(columns={A_BasketGame0316.columns[1]:"BasketGame0316"}, inplace =True)
        A_BasketGame0316['BasketGame0316']=A_BasketGame0316['BasketGame0316'] * 60
        
        '''
        A_BasketGame0317
        '''
        A_BasketGame0317 = pd.read_excel('D:\Mass\Sports_data\Active\Comp\A_BasketGame0317.xlsx')
        A_BasketGame0317 =A_BasketGame0317.drop_duplicates(subset = None, keep = 'last')
        A_BasketGame0317.rename(columns={A_BasketGame0317.columns[0]: "Student_ID" }, inplace=True)
        A_BasketGame0317.rename(columns={A_BasketGame0317.columns[1]:"BasketGame0317"}, inplace =True)
        A_BasketGame0317['BasketGame0317']=A_BasketGame0317['BasketGame0317'] * 60
        
        '''
        A_Swim0119
        '''
        A_Swim0219 =pd.read_excel('D:\Mass\Sports_data\Active\Comp\A_Swim0119.xlsx')
        A_Swim0219 =A_Swim0219.drop_duplicates(subset = None, keep = 'last')
        A_Swim0219.rename(columns={A_Swim0219.columns[0]:"Student_ID"}, inplace=True)
        A_Swim0219.rename(columns={A_Swim0219.columns[1]:"Swim_In0119"}, inplace=True )
        A_Swim0219.rename(columns={A_Swim0219.columns[2]:"Swim_Out0119"},inplace=True)
        
        A_Swim0219['Swim0119']=A_Swim0219['Swim_Out0119']-A_Swim0219['Swim_In0119']
        A_Swim0219['Swim0119']=A_Swim0219['Swim0119'] / np.timedelta64(60,'s')
        
        '''
        A_Swim0126
        '''
        A_Swim0226 =pd.read_excel('D:\Mass\Sports_data\Active\Comp\A_Swim0126.xlsx')
        A_Swim0226 =A_Swim0226.drop_duplicates(subset = None, keep = 'last')
        A_Swim0226.rename(columns={A_Swim0226.columns[0]:"Student_ID"}, inplace=True)
        A_Swim0226.rename(columns={A_Swim0226.columns[1]:"Swim_In0126"}, inplace=True )
        A_Swim0226.rename(columns={A_Swim0226.columns[2]:"Swim_Out0126"},inplace=True)
        
        A_Swim0226['Swim0126']=A_Swim0226['Swim_Out0126']-A_Swim0226['Swim_In0126']
        A_Swim0226['Swim0126']=A_Swim0226['Swim0126'] / np.timedelta64(60,'s')
        
        
        '''
        A_Soccer0330
        '''
        A_Soccer0330 =pd.read_excel('D:\Mass\Sports_data\Active\Comp\A_Soccer0330.xlsx')
        A_Soccer0330 =A_Soccer0330.drop_duplicates(subset = None, keep = 'last')
        A_Soccer0330.rename(columns={A_Soccer0330.columns[0]:"Student_ID"}, inplace=True)
        A_Soccer0330.rename(columns={A_Soccer0330.columns[1]:"Soccer0330_In"}, inplace=True )
        A_Soccer0330.rename(columns={A_Soccer0330.columns[2]:"Soccer0330_Out"},inplace=True)
        
        A_Soccer0330['Soccer0330']=A_Soccer0330['Soccer0330_Out']-A_Soccer0330['Soccer0330_In']
        A_Soccer0330['Soccer0330']=A_Soccer0330['Soccer0330'] / np.timedelta64(60,'s')
        
        '''
        A_Soccer0331
        '''
        A_Soccer0331 =pd.read_excel('D:\Mass\Sports_data\Active\Comp\A_Soccer0331.xlsx')
        A_Soccer0331 =A_Soccer0331.drop_duplicates(subset = None, keep = 'last')
        A_Soccer0331.rename(columns={A_Soccer0331.columns[0]:"Student_ID"}, inplace=True)
        A_Soccer0331.rename(columns={A_Soccer0331.columns[1]:"Soccer0331_In"}, inplace=True )
        A_Soccer0331.rename(columns={A_Soccer0331.columns[2]:"Soccer0331_Out"},inplace=True)
        
        A_Soccer0331['Soccer0331']=A_Soccer0331['Soccer0331_Out']-A_Soccer0331['Soccer0331_In']
        A_Soccer0331['Soccer0331']=A_Soccer0331['Soccer0331'] / np.timedelta64(60,'s')
        
        '''
        A_Squash0330
        '''
        A_Squash0330 =pd.read_excel('D:\Mass\Sports_data\Active\Comp\A_Squash0330.xlsx')
        A_Squash0330 =A_Squash0330.drop_duplicates(subset = None, keep = 'last')
        A_Squash0330.rename(columns={A_Squash0330.columns[0]:"Student_ID"}, inplace=True)
        A_Squash0330.rename(columns={A_Squash0330.columns[1]:"Squash0330_In"}, inplace=True )
        A_Squash0330.rename(columns={A_Squash0330.columns[2]:"Squash0330_Out"},inplace=True)
        
        A_Squash0330['Squash_Competition']=A_Squash0330['Squash0330_Out']-A_Squash0330['Squash0330_In']
        A_Squash0330['Squash_Competition']=A_Squash0330['Squash_Competition'] / np.timedelta64(60,'s')
        
        
        '''
        A_Badminton
        '''
        A_Badminton = pd.read_excel('D:\Mass\Sports_data\Active\Comp\A_Badminton.xlsx')
        
        # #活动信息清洗
        A_Badminton =A_Badminton.drop_duplicates(subset = None, keep = 'last')
        #列名更改
        A_Badminton.rename(columns={A_Badminton.columns[0]: "Student_ID" }, inplace=True)
        A_Badminton.rename(columns={A_Badminton.columns[1]:"Badminton_In"}, inplace =True)
        A_Badminton.rename(columns={A_Badminton.columns[2]:"Badminton_Out"}, inplace =True)
        
        #时间差值计算
        A_Badminton['BadmintonCompetition']=A_Badminton['Badminton_Out']-A_Badminton['Badminton_In']
        #转换为float型的分钟
        A_Badminton['BadmintonCompetition']=A_Badminton['BadmintonCompetition'] / np.timedelta64(60,'s')
        
        
        '''
        A_TableTennis
        '''
        A_TableTennis = pd.read_excel('D:\Mass\Sports_data\Active\Comp\A_TableTennis.xlsx')
        A_TableTennis =A_TableTennis.drop_duplicates(subset = None, keep = 'last')
        A_TableTennis.rename(columns={A_TableTennis.columns[0]: "Student_ID" }, inplace=True)
        A_TableTennis.rename(columns={A_TableTennis.columns[1]:"Tennis_In"}, inplace =True)
        A_TableTennis.rename(columns={A_TableTennis.columns[2]:"Tennis_Out"}, inplace =True)
        
        A_TableTennis['TableTennisCompetition']=A_TableTennis['Tennis_Out']-A_TableTennis['Tennis_In']
        A_TableTennis['TableTennisCompetition']=A_TableTennis['TableTennisCompetition'] / np.timedelta64(60,'s')
        
        '''
        A_TableTennisClub
        '''
        TableTennisClub = pd.read_excel('D:\Mass\Sports_data\Active\Comp\A_TableTennisClub.xlsx')
        TableTennisClub =TableTennisClub.drop_duplicates(subset = None, keep = 'last')
        TableTennisClub.rename(columns={TableTennisClub.columns[0]: "Student_ID" }, inplace=True)
        TableTennisClub.rename(columns={TableTennisClub.columns[1]:"TableTennisClubCompetition"}, inplace =True)
        TableTennisClub['TableTennisClubCompetition']=TableTennisClub['TableTennisClubCompetition'] * 60
        
        
        
        
        '''
        A_Track&Field
        '''
        A_TrackField = pd.read_excel('D:\Mass\Sports_data\Active\Comp\A_TrackField.xlsx')
        A_TrackField =A_TrackField.drop_duplicates(subset = None, keep = 'last')
        A_TrackField.rename(columns={A_TrackField.columns[0]: "Student_ID" }, inplace=True)
        A_TrackField.rename(columns={A_TrackField.columns[1]:"TrackField_In"}, inplace =True)
        A_TrackField.rename(columns={A_TrackField.columns[2]:"TrackField_Out"}, inplace =True)
        
        A_TrackField['TrackFieldCompetition']=A_TrackField['TrackField_Out']-A_TrackField['TrackField_In']
        A_TrackField['TrackFieldCompetition']=A_TrackField['TrackFieldCompetition'] / np.timedelta64(60,'s')
        
        
        '''
        A_Volleyball
        '''
        A_Volleyball = pd.read_excel('D:\Mass\Sports_data\Active\Comp\A_Volleyball.xlsx')
        A_Volleyball =A_Volleyball.drop_duplicates(subset = None, keep = 'last')
        A_Volleyball.rename(columns={A_Volleyball.columns[0]: "Student_ID" }, inplace=True)
        A_Volleyball.rename(columns={A_Volleyball.columns[1]:"Volleyball_In"}, inplace =True)
        A_Volleyball.rename(columns={A_Volleyball.columns[2]:"Volleyball_Out"}, inplace =True)
        
        A_Volleyball['A_Volleyball']=A_Volleyball['Volleyball_Out']-A_Volleyball['Volleyball_In']
        A_Volleyball['A_Volleyball']=A_Volleyball['A_Volleyball'] / np.timedelta64(60,'s')
        
        '''
        A_Volleyball0323
        '''
        A_Volleyball0323 = pd.read_excel('D:\Mass\Sports_data\Active\Comp\A_Volleyball0323.xlsx')
        A_Volleyball0323 =A_Volleyball0323.drop_duplicates(subset = None, keep = 'last')
        A_Volleyball0323.rename(columns={A_Volleyball0323.columns[0]: "Student_ID" }, inplace=True)
        A_Volleyball0323.rename(columns={A_Volleyball0323.columns[1]:"Volleyball0323_In"}, inplace =True)
        A_Volleyball0323.rename(columns={A_Volleyball0323.columns[2]:"Volleyball0323_Out"}, inplace =True)
        
        A_Volleyball0323['A_Volleyball0323']=A_Volleyball0323['Volleyball0323_Out']-A_Volleyball0323['Volleyball0323_In']
        A_Volleyball0323['A_Volleyball0323']=A_Volleyball0323['A_Volleyball0323'] / np.timedelta64(60,'s')
        
        
        '''
        A_Volleyball0323PM
        '''
        A_Volleyball0323PM = pd.read_excel('D:\Mass\Sports_data\Active\Comp\A_Volleyball0323PM.xlsx')
        A_Volleyball0323PM =A_Volleyball0323PM.drop_duplicates(subset = None, keep = 'last')
        A_Volleyball0323PM.rename(columns={A_Volleyball0323PM.columns[0]: "Student_ID" }, inplace=True)
        A_Volleyball0323PM.rename(columns={A_Volleyball0323PM.columns[1]:"Volleyball0323PM_In"}, inplace =True)
        A_Volleyball0323PM.rename(columns={A_Volleyball0323PM.columns[2]:"Volleyball0323PM_Out"}, inplace =True)
        
        A_Volleyball0323PM['A_Volleyball0323PM']=A_Volleyball0323PM['Volleyball0323PM_Out']-A_Volleyball0323PM['Volleyball0323PM_In']
        A_Volleyball0323PM['A_Volleyball0323PM']=A_Volleyball0323PM['A_Volleyball0323PM'] / np.timedelta64(60,'s')
        
        '''
        A_Volleyball0324
        '''
        A_Volleyball0324 = pd.read_excel('D:\Mass\Sports_data\Active\Comp\A_Volleyball0324.xlsx')
        A_Volleyball0324 =A_Volleyball0324.drop_duplicates(subset = None, keep = 'last')
        A_Volleyball0324.rename(columns={A_Volleyball0324.columns[0]: "Student_ID" }, inplace=True)
        A_Volleyball0324.rename(columns={A_Volleyball0324.columns[1]:"Volleyball0324_In"}, inplace =True)
        A_Volleyball0324.rename(columns={A_Volleyball0324.columns[2]:"Volleyball0324_Out"}, inplace =True)
        
        A_Volleyball0324['A_Volleyball0324']=A_Volleyball0324['Volleyball0324_Out']-A_Volleyball0324['Volleyball0324_In']
        A_Volleyball0324['A_Volleyball0324']=A_Volleyball0324['A_Volleyball0324'] / np.timedelta64(60,'s')
        
        '''
        A_UT_07Badminton_University
        '''
        Badminton_Un = pd.read_excel('D:\Mass\Sports_data\Active\Comp\A_UT_07Badminton_University.xlsx')
        Badminton_Un =Badminton_Un.drop_duplicates(subset = None, keep = 'last')
        Badminton_Un.rename(columns={Badminton_Un.columns[0]: "Student_ID" }, inplace=True)
        Badminton_Un.rename(columns={Badminton_Un.columns[1]:"Badminton_Un_Competition"}, inplace =True)
        Badminton_Un['Badminton_Un_Competition']=Badminton_Un['Badminton_Un_Competition'] * 60
        
        '''
        A_UT_10DragonDance
        '''
        DragonDance = pd.read_excel('D:\Mass\Sports_data\Active\Comp\A_UT_10DragonDance.xlsx')
        DragonDance =DragonDance.drop_duplicates(subset = None, keep = 'last')
        DragonDance.rename(columns={DragonDance.columns[0]: "Student_ID" }, inplace=True)
        DragonDance.rename(columns={DragonDance.columns[1]:"DragonDance"}, inplace =True)
        DragonDance['DragonDance']=DragonDance['DragonDance'] * 60
        
        
        '''
        A_UT_11BasketballMen
        '''
        BasketballMen = pd.read_excel('D:\Mass\Sports_data\Active\Comp\A_UT_11BasketballMen.xlsx')
        BasketballMen =BasketballMen.drop_duplicates(subset = None, keep = 'last')
        BasketballMen.rename(columns={BasketballMen.columns[0]: "Student_ID" }, inplace=True)
        BasketballMen.rename(columns={BasketballMen.columns[1]:"BasketballMen"}, inplace =True)
        BasketballMen['BasketballMen']=BasketballMen['BasketballMen'] * 60
        
        '''
        A_UT_12BasketballMenCD
        '''
        BasketballMenCD = pd.read_excel('D:\Mass\Sports_data\Active\Comp\A_UT_12BasketballMenCD.xlsx')
        BasketballMenCD =BasketballMenCD.drop_duplicates(subset = None, keep = 'last')
        BasketballMenCD.rename(columns={BasketballMenCD.columns[0]: "Student_ID" }, inplace=True)
        BasketballMenCD.rename(columns={BasketballMenCD.columns[1]:"BasketballMenCD"}, inplace =True)
        BasketballMenCD['BasketballMenCD']=BasketballMenCD['BasketballMenCD'] * 60
        
        '''
        A_UT_13BasketballWomenCD
        '''
        BasketballWomenCD = pd.read_excel('D:\Mass\Sports_data\Active\Comp\A_UT_13BasketballWomenCD.xlsx')
        BasketballWomenCD =BasketballWomenCD.drop_duplicates(subset = None, keep = 'last')
        BasketballWomenCD.rename(columns={BasketballWomenCD.columns[0]: "Student_ID" }, inplace=True)
        BasketballWomenCD.rename(columns={BasketballWomenCD.columns[1]:"BasketballWomenCD"}, inplace =True)
        BasketballWomenCD['BasketballWomenCD']=BasketballWomenCD['BasketballWomenCD'] * 60
        
        '''
        A_UT_14BasketballWomen
        '''
        BasketballWomen = pd.read_excel('D:\Mass\Sports_data\Active\Comp\A_UT_14BasketballWomen.xlsx')
        BasketballWomen =BasketballWomen.drop_duplicates(subset = None, keep = 'last')
        BasketballWomen.rename(columns={BasketballWomen.columns[0]: "Student_ID" }, inplace=True)
        BasketballWomen.rename(columns={BasketballWomen.columns[1]:"BasketballWomen"}, inplace =True)
        BasketballWomen['BasketballWomen']=BasketballWomen['BasketballWomen'] * 60
        
        '''
        A_UT_15Fencing
        '''
        Fencing = pd.read_excel('D:\Mass\Sports_data\Active\Comp\A_UT_15Fencing.xlsx')
        Fencing =Fencing.drop_duplicates(subset = None, keep = 'last')
        Fencing.rename(columns={Fencing.columns[0]: "Student_ID" }, inplace=True)
        Fencing.rename(columns={Fencing.columns[1]:"FencingCompetition"}, inplace =True)
        Fencing['FencingCompetition']=Fencing['FencingCompetition'] * 60
        
        
        '''
        A_UT_16Opening0427
        '''
        Opening0427 = pd.read_excel('D:\Mass\Sports_data\Active\Comp\A_UT_16Opening0427.xlsx')
        Opening0427 =Opening0427.drop_duplicates(subset = None, keep = 'last')
        Opening0427.rename(columns={Opening0427.columns[0]: "Student_ID" }, inplace=True)
        Opening0427.rename(columns={Opening0427.columns[1]:"Opening_AoKeBei0427"}, inplace =True)
        Opening0427['Opening_AoKeBei0427']=Opening0427['Opening_AoKeBei0427'] * 60
        
        '''
        A_UT_17Competi0427
        '''
        Competi0427 = pd.read_excel('D:\Mass\Sports_data\Active\Comp\A_UT_17Competi0427.xlsx')
        Competi0427 =Competi0427.drop_duplicates(subset = None, keep = 'last')
        Competi0427.rename(columns={Competi0427.columns[0]: "Student_ID" }, inplace=True)
        Competi0427.rename(columns={Competi0427.columns[1]:"Competi_AoKeBei0427"}, inplace =True)
        Competi0427['Competi_AoKeBei0427']=Competi0427['Competi_AoKeBei0427'] * 60
        
        '''
        A_UT_18Dancing0413
        '''
        Dancing0413 = pd.read_excel('D:\Mass\Sports_data\Active\Comp\A_UT_18Dancing0413.xlsx')
        Dancing0413 =Dancing0413.drop_duplicates(subset = None, keep = 'last')
        Dancing0413.rename(columns={Dancing0413.columns[0]: "Student_ID" }, inplace=True)
        Dancing0413.rename(columns={Dancing0413.columns[1]:"DancingTongXin"}, inplace =True)
        Dancing0413['DancingTongXin']=Dancing0413['DancingTongXin'] * 60
        
        '''
        A_UT_19Dancing0424
        '''
        Dancing0424 = pd.read_excel('D:\Mass\Sports_data\Active\Comp\A_UT_19Dancing0424.xlsx')
        Dancing0424 =Dancing0424.drop_duplicates(subset = None, keep = 'last')
        Dancing0424.rename(columns={Dancing0424.columns[0]: "Student_ID" }, inplace=True)
        Dancing0424.rename(columns={Dancing0424.columns[1]:"DancingNianLun"}, inplace =True)
        Dancing0424['DancingNianLun']=Dancing0424['DancingNianLun'] * 60
        
        '''
        A_UT_20FlagRaising
        '''
        FlagRaising = pd.read_excel('D:\Mass\Sports_data\Active\Comp\A_UT_20FlagRaising.xlsx')
        FlagRaising =FlagRaising.drop_duplicates(subset = None, keep = 'last')
        FlagRaising.rename(columns={FlagRaising.columns[0]: "Student_ID" }, inplace=True)
        FlagRaising.rename(columns={FlagRaising.columns[1]:"FlagRaising"}, inplace =True)
        FlagRaising['FlagRaising']=FlagRaising['FlagRaising'] * 60
        
        
        '''
        A_UT_21Flag0430
        '''
        Flag0430 = pd.read_excel('D:\Mass\Sports_data\Active\Comp\A_UT_21Flag0430.xlsx')
        Flag0430 =Flag0430.drop_duplicates(subset = None, keep = 'last')
        Flag0430.rename(columns={Flag0430.columns[0]: "Student_ID" }, inplace=True)
        Flag0430.rename(columns={Flag0430.columns[1]:"Flag54Ceremony"}, inplace =True)
        Flag0430['Flag54Ceremony']=Flag0430['Flag54Ceremony'] * 60

        
        '''
        A_Climbing0309_session1
        '''
        ClimbingSession1 = pd.read_excel('D:\Mass\Sports_data\Active\Comp\A_Climbing0309_session1.xlsx')
        ClimbingSession1 =ClimbingSession1.drop_duplicates(subset = None, keep = 'last')
        ClimbingSession1.rename(columns={ClimbingSession1.columns[0]: "Student_ID" }, inplace=True)
        ClimbingSession1.rename(columns={ClimbingSession1.columns[1]:"ClimbingSession1"}, inplace =True)
        ClimbingSession1['ClimbingSession1']=ClimbingSession1['ClimbingSession1'] * 60
        
        
        '''
        A_Climbing0309_session2
        '''
        ClimbingSession2 = pd.read_excel('D:\Mass\Sports_data\Active\Comp\A_Climbing0309_session2.xlsx')
        ClimbingSession2 =ClimbingSession2.drop_duplicates(subset = None, keep = 'last')
        ClimbingSession2.rename(columns={ClimbingSession2.columns[0]: "Student_ID" }, inplace=True)
        ClimbingSession2.rename(columns={ClimbingSession2.columns[1]:"ClimbingSession2"}, inplace =True)
        ClimbingSession2['ClimbingSession2']=ClimbingSession2['ClimbingSession2'] * 60

        '''
        A_HKUST_UM_Cup0427
        '''
        HKUST_UM_Cup = pd.read_excel('D:\Mass\Sports_data\Active\Comp\A_HKUST_UM_Cup0427.xlsx')
        HKUST_UM_Cup =HKUST_UM_Cup.drop_duplicates(subset = None, keep = 'last')
        HKUST_UM_Cup.rename(columns={HKUST_UM_Cup.columns[0]: "Student_ID" }, inplace=True)
        HKUST_UM_Cup.rename(columns={HKUST_UM_Cup.columns[1]:"HKUST_UM_Cup"}, inplace =True)
        HKUST_UM_Cup['HKUST_UM_Cup']=HKUST_UM_Cup['HKUST_UM_Cup'] * 60

        '''
        A_FSTBball
        '''
        FSTBball = pd.read_excel('D:\Mass\Sports_data\Active\Comp\A_FSTBball.xlsx')
        FSTBball =FSTBball.drop_duplicates(subset = None, keep = 'last')
        FSTBball.rename(columns={FSTBball.columns[0]: "Student_ID" }, inplace=True)
        FSTBball.rename(columns={FSTBball.columns[1]:"FSTBball"}, inplace =True)
        FSTBball['FSTBball']=FSTBball['FSTBball'] * 60

        '''
        A_FSTFutsal
        '''
        FSTFutsal = pd.read_excel('D:\Mass\Sports_data\Active\Comp\A_FSTFutsal.xlsx')
        FSTFutsal =FSTFutsal.drop_duplicates(subset = None, keep = 'last')
        FSTFutsal.rename(columns={FSTFutsal.columns[0]: "Student_ID" }, inplace=True)
        FSTFutsal.rename(columns={FSTFutsal.columns[1]:"FSTFutsal"}, inplace =True)
        FSTFutsal['FSTFutsal']=FSTFutsal['FSTFutsal'] * 60
        
        
        '''
        A_UMOpening
        '''
        A_UMOpening = pd.read_excel('D:\Mass\Sports_data\Active\Comp\A_UMOpening.xlsx')
        A_UMOpening =A_UMOpening.drop_duplicates(subset = None, keep = 'last')
        A_UMOpening.rename(columns={A_UMOpening.columns[0]: "Student_ID" }, inplace=True)
        A_UMOpening.rename(columns={A_UMOpening.columns[1]:"UMOpening_In"}, inplace =True)
        A_UMOpening.rename(columns={A_UMOpening.columns[2]:"UMOpening_Out"}, inplace =True)
        
        A_UMOpening['Opening_UM']=A_UMOpening['UMOpening_Out']-A_UMOpening['UMOpening_In']
        A_UMOpening['Opening_UM']=A_UMOpening['Opening_UM'] / np.timedelta64(60,'s')
        
        
        '''
        A_JunKR20190108
        '''
        A_JunKR0108 = pd.read_excel('D:\Mass\Sports_data\Active\Runing\A_JanuRun0108.xlsx')
        
        # #活动信息清洗
        A_JunKR0108 =A_JunKR0108.drop_duplicates(subset = None, keep = 'last')
        #列名更改
        A_JunKR0108.rename(columns={A_JunKR0108.columns[0]: "Student_ID" }, inplace=True)
        A_JunKR0108.rename(columns={A_JunKR0108.columns[1]:"JunKR0108_In"}, inplace =True)
        A_JunKR0108.rename(columns={A_JunKR0108.columns[2]:"JunKR0108_Out"}, inplace =True)
        
        #时间差值计算
        A_JunKR0108['KeepRun0108']=A_JunKR0108['JunKR0108_Out']-A_JunKR0108['JunKR0108_In']
        #转换为float型的分钟
        A_JunKR0108['KeepRun0108']=A_JunKR0108['KeepRun0108'] / np.timedelta64(60,'s')
        
        '''
        A_JunKR20190110
        '''
        A_JunKR0110 = pd.read_excel('D:\Mass\Sports_data\Active\Runing\A_JanuRun0110.xlsx')
        A_JunKR0110 =A_JunKR0110.drop_duplicates(subset = None, keep = 'last')
        A_JunKR0110.rename(columns={A_JunKR0110.columns[0]: "Student_ID" }, inplace=True)
        A_JunKR0110.rename(columns={A_JunKR0110.columns[1]:"JunKR0110_In"}, inplace =True)
        A_JunKR0110.rename(columns={A_JunKR0110.columns[2]:"JunKR0110_Out"}, inplace =True)
        
        A_JunKR0110['KeepRun0110']=A_JunKR0110['JunKR0110_Out']-A_JunKR0110['JunKR0110_In']
        A_JunKR0110['KeepRun0110']=A_JunKR0110['KeepRun0110'] / np.timedelta64(60,'s')
        
        '''
        A_JunKR20190117
        '''
        A_JunKR0117 = pd.read_excel('D:\Mass\Sports_data\Active\Runing\A_JanuRun0117.xlsx')
        A_JunKR0117 =A_JunKR0117.drop_duplicates(subset = None, keep = 'last')
        A_JunKR0117.rename(columns={A_JunKR0117.columns[0]: "Student_ID" }, inplace=True)
        A_JunKR0117.rename(columns={A_JunKR0117.columns[1]:"JunKR0117_In"}, inplace =True)
        A_JunKR0117.rename(columns={A_JunKR0117.columns[2]:"JunKR0117_Out"}, inplace =True)
        
        A_JunKR0117['KeepRun0117']=A_JunKR0117['JunKR0117_Out']-A_JunKR0117['JunKR0117_In']
        A_JunKR0117['KeepRun0117']=A_JunKR0117['KeepRun0117'] / np.timedelta64(60,'s')
        
        '''
        A_JunKR20190122
        '''
        A_JunKR0122 = pd.read_excel('D:\Mass\Sports_data\Active\Runing\A_JanuRun0122.xlsx')
        A_JunKR0122 =A_JunKR0122.drop_duplicates(subset = None, keep = 'last')
        A_JunKR0122.rename(columns={A_JunKR0122.columns[0]: "Student_ID" }, inplace=True)
        A_JunKR0122.rename(columns={A_JunKR0122.columns[1]:"JunKR0122_In"}, inplace =True)
        A_JunKR0122.rename(columns={A_JunKR0122.columns[2]:"JunKR0122_Out"}, inplace =True)
        
        A_JunKR0122['KeepRun0122']=A_JunKR0122['JunKR0122_Out']-A_JunKR0122['JunKR0122_In']
        A_JunKR0122['KeepRun0122']=A_JunKR0122['KeepRun0122'] / np.timedelta64(60,'s')
        
        '''
        A_JunKR20190124
        '''
        A_JunKR0124 = pd.read_excel('D:\Mass\Sports_data\Active\Runing\A_JanuRun0124.xlsx')
        A_JunKR0124 =A_JunKR0124.drop_duplicates(subset = None, keep = 'last')
        A_JunKR0124.rename(columns={A_JunKR0124.columns[0]: "Student_ID" }, inplace=True)
        A_JunKR0124.rename(columns={A_JunKR0124.columns[1]:"JunKR0124_In"}, inplace =True)
        A_JunKR0124.rename(columns={A_JunKR0124.columns[2]:"JunKR0124_Out"}, inplace =True)
        
        A_JunKR0124['KeepRun0124']=A_JunKR0124['JunKR0124_Out']-A_JunKR0124['JunKR0124_In']
        A_JunKR0124['KeepRun0124']=A_JunKR0124['KeepRun0124'] / np.timedelta64(60,'s')
        
        '''
        A_JunKR20190129
        '''
        A_JunKR0129 = pd.read_excel('D:\Mass\Sports_data\Active\Runing\A_JanuRun0129.xlsx')
        A_JunKR0129 =A_JunKR0129.drop_duplicates(subset = None, keep = 'last')
        A_JunKR0129.rename(columns={A_JunKR0129.columns[0]: "Student_ID" }, inplace=True)
        A_JunKR0129.rename(columns={A_JunKR0129.columns[1]:"JunKR0129_In"}, inplace =True)
        A_JunKR0129.rename(columns={A_JunKR0129.columns[2]:"JunKR0129_Out"}, inplace =True)
        
        A_JunKR0129['KeepRun0129']=A_JunKR0129['JunKR0129_Out']-A_JunKR0129['JunKR0129_In']
        A_JunKR0129['KeepRun0129']=A_JunKR0129['KeepRun0129'] / np.timedelta64(60,'s')
        
        '''
        A_NightRun20190117
        '''
        A_NightRun0117 =pd.read_excel('D:\Mass\Sports_data\Active\Runing\A_NightRun20190117.xlsx')
        A_NightRun0117 =A_NightRun0117.drop_duplicates(subset = None, keep = 'last')
        A_NightRun0117.rename(columns={A_NightRun0117.columns[0]:"Student_ID"}, inplace=True)
        A_NightRun0117.rename(columns={A_NightRun0117.columns[1]:"NightRun_In0117"}, inplace=True )
        A_NightRun0117.rename(columns={A_NightRun0117.columns[2]:"NightRun_Out0117"},inplace=True)
        
        A_NightRun0117['NightRun0117']=A_NightRun0117['NightRun_Out0117']-A_NightRun0117['NightRun_In0117']
        A_NightRun0117['NightRun0117']=A_NightRun0117['NightRun0117'] / np.timedelta64(60,'s')
        
        '''
        A_KeepRun0305
        '''
        A_KeepRun0305 = pd.read_excel('D:\Mass\Sports_data\Active\Runing\A_KeepRun0305.xlsx')
        A_KeepRun0305 =A_KeepRun0305.drop_duplicates(subset = None, keep = 'last')
        A_KeepRun0305.rename(columns={A_KeepRun0305.columns[0]: "Student_ID" }, inplace=True)
        A_KeepRun0305.rename(columns={A_KeepRun0305.columns[1]:"KeepRun0305_In"}, inplace =True)
        A_KeepRun0305.rename(columns={A_KeepRun0305.columns[2]:"KeepRun0305_Out"}, inplace =True)
        
        A_KeepRun0305['KeepRun0305']=A_KeepRun0305['KeepRun0305_Out']-A_KeepRun0305['KeepRun0305_In']
        A_KeepRun0305['KeepRun0305']=A_KeepRun0305['KeepRun0305'] / np.timedelta64(60,'s')
        
        '''
        A_KeepRun0312AM
        '''
        A_KeepRun0312AM = pd.read_excel('D:\Mass\Sports_data\Active\Runing\A_KeepRun0312AM.xlsx')
        A_KeepRun0312AM =A_KeepRun0312AM.drop_duplicates(subset = None, keep = 'last')
        A_KeepRun0312AM.rename(columns={A_KeepRun0312AM.columns[0]: "Student_ID" }, inplace=True)
        A_KeepRun0312AM.rename(columns={A_KeepRun0312AM.columns[1]:"KeepRun0312AM_In"}, inplace =True)
        A_KeepRun0312AM.rename(columns={A_KeepRun0312AM.columns[2]:"KeepRun0312AM_Out"}, inplace =True)
        
        A_KeepRun0312AM['KeepRun0312AM']=A_KeepRun0312AM['KeepRun0312AM_Out']-A_KeepRun0312AM['KeepRun0312AM_In']
        A_KeepRun0312AM['KeepRun0312AM']=A_KeepRun0312AM['KeepRun0312AM'] / np.timedelta64(60,'s')
        
        '''
        A_KeepRun0312PM
        '''
        A_KeepRun0312PM = pd.read_excel('D:\Mass\Sports_data\Active\Runing\A_KeepRun0312PM.xlsx')
        A_KeepRun0312PM =A_KeepRun0312PM.drop_duplicates(subset = None, keep = 'last')
        A_KeepRun0312PM.rename(columns={A_KeepRun0312PM.columns[0]: "Student_ID" }, inplace=True)
        A_KeepRun0312PM.rename(columns={A_KeepRun0312PM.columns[1]:"KeepRun0312PM_In"}, inplace =True)
        A_KeepRun0312PM.rename(columns={A_KeepRun0312PM.columns[2]:"KeepRun0312PM_Out"}, inplace =True)
        
        A_KeepRun0312PM['KeepRun0312PM']=A_KeepRun0312PM['KeepRun0312PM_Out']-A_KeepRun0312PM['KeepRun0312PM_In']
        A_KeepRun0312PM['KeepRun0312PM']=A_KeepRun0312PM['KeepRun0312PM'] / np.timedelta64(60,'s')
        
        
        '''
        A_KeepRun0313
        '''
        A_KeepRun0313 = pd.read_excel('D:\Mass\Sports_data\Active\Runing\A_KeepRun0313.xlsx')
        A_KeepRun0313 =A_KeepRun0313.drop_duplicates(subset = None, keep = 'last')
        A_KeepRun0313.rename(columns={A_KeepRun0313.columns[0]: "Student_ID" }, inplace=True)
        A_KeepRun0313.rename(columns={A_KeepRun0313.columns[1]:"KeepRun0313_In"}, inplace =True)
        A_KeepRun0313.rename(columns={A_KeepRun0313.columns[2]:"KeepRun0313_Out"}, inplace =True)
        
        A_KeepRun0313['KeepRun0313']=A_KeepRun0313['KeepRun0313_Out']-A_KeepRun0313['KeepRun0313_In']
        A_KeepRun0313['KeepRun0313']=A_KeepRun0313['KeepRun0313'] / np.timedelta64(60,'s')
        
        '''
        A_KeepRun0314
        '''
        A_KeepRun0314 = pd.read_excel('D:\Mass\Sports_data\Active\Runing\A_KeepRun0314.xlsx')
        A_KeepRun0314 =A_KeepRun0314.drop_duplicates(subset = None, keep = 'last')
        A_KeepRun0314.rename(columns={A_KeepRun0314.columns[0]: "Student_ID" }, inplace=True)
        A_KeepRun0314.rename(columns={A_KeepRun0314.columns[1]:"KeepRun0314_In"}, inplace =True)
        A_KeepRun0314.rename(columns={A_KeepRun0314.columns[2]:"KeepRun0314_Out"}, inplace =True)
        
        A_KeepRun0314['KeepRun0314']=A_KeepRun0314['KeepRun0314_Out']-A_KeepRun0314['KeepRun0314_In']
        A_KeepRun0314['KeepRun0314']=A_KeepRun0314['KeepRun0314'] / np.timedelta64(60,'s')
        
        '''
        A_KeepRun0319AM
        '''
        A_KeepRun0319AM = pd.read_excel('D:\Mass\Sports_data\Active\Runing\A_KeepRun0319AM.xlsx')
        A_KeepRun0319AM =A_KeepRun0319AM.drop_duplicates(subset = None, keep = 'last')
        A_KeepRun0319AM.rename(columns={A_KeepRun0319AM.columns[0]: "Student_ID" }, inplace=True)
        A_KeepRun0319AM.rename(columns={A_KeepRun0319AM.columns[1]:"KeepRun0319AM_In"}, inplace =True)
        A_KeepRun0319AM.rename(columns={A_KeepRun0319AM.columns[2]:"KeepRun0319AM_Out"}, inplace =True)
        
        A_KeepRun0319AM['KeepRun0319AM']=A_KeepRun0319AM['KeepRun0319AM_Out']-A_KeepRun0319AM['KeepRun0319AM_In']
        A_KeepRun0319AM['KeepRun0319AM']=A_KeepRun0319AM['KeepRun0319AM'] / np.timedelta64(60,'s')
        
        '''
        A_KeepRun0319PM
        '''
        A_KeepRun0319PM = pd.read_excel('D:\Mass\Sports_data\Active\Runing\A_KeepRun0319PM.xlsx')
        A_KeepRun0319PM =A_KeepRun0319PM.drop_duplicates(subset = None, keep = 'last')
        A_KeepRun0319PM.rename(columns={A_KeepRun0319PM.columns[0]: "Student_ID" }, inplace=True)
        A_KeepRun0319PM.rename(columns={A_KeepRun0319PM.columns[1]:"KeepRun0319PM_In"}, inplace =True)
        A_KeepRun0319PM.rename(columns={A_KeepRun0319PM.columns[2]:"KeepRun0319PM_Out"}, inplace =True)
        
        A_KeepRun0319PM['KeepRun0319PM']=A_KeepRun0319PM['KeepRun0319PM_Out']-A_KeepRun0319PM['KeepRun0319PM_In']
        A_KeepRun0319PM['KeepRun0319PM']=A_KeepRun0319PM['KeepRun0319PM'] / np.timedelta64(60,'s')
        
        '''
        A_KeepRun0320
        '''
        A_KeepRun0320 = pd.read_excel('D:\Mass\Sports_data\Active\Runing\A_KeepRun0320.xlsx')
        A_KeepRun0320 =A_KeepRun0320.drop_duplicates(subset = None, keep = 'last')
        A_KeepRun0320.rename(columns={A_KeepRun0320.columns[0]: "Student_ID" }, inplace=True)
        A_KeepRun0320.rename(columns={A_KeepRun0320.columns[1]:"KeepRun0320_In"}, inplace =True)
        A_KeepRun0320.rename(columns={A_KeepRun0320.columns[2]:"KeepRun0320_Out"}, inplace =True)
        
        A_KeepRun0320['KeepRun0320']=A_KeepRun0320['KeepRun0320_Out']-A_KeepRun0320['KeepRun0320_In']
        A_KeepRun0320['KeepRun0320']=A_KeepRun0320['KeepRun0320'] / np.timedelta64(60,'s')
        
        '''
        A_KeepRun0321
        '''
        A_KeepRun0321 = pd.read_excel('D:\Mass\Sports_data\Active\Runing\A_KeepRun0321.xlsx')
        A_KeepRun0321 =A_KeepRun0321.drop_duplicates(subset = None, keep = 'last')
        A_KeepRun0321.rename(columns={A_KeepRun0321.columns[0]: "Student_ID" }, inplace=True)
        A_KeepRun0321.rename(columns={A_KeepRun0321.columns[1]:"KeepRun0321_In"}, inplace =True)
        A_KeepRun0321.rename(columns={A_KeepRun0321.columns[2]:"KeepRun0321_Out"}, inplace =True)
        
        A_KeepRun0321['KeepRun0321']=A_KeepRun0321['KeepRun0321_Out']-A_KeepRun0321['KeepRun0321_In']
        A_KeepRun0321['KeepRun0321']=A_KeepRun0321['KeepRun0321'] / np.timedelta64(60,'s')
        
        '''
        A_KeepRun0326AM
        '''
        A_KeepRun0326AM = pd.read_excel('D:\Mass\Sports_data\Active\Runing\A_KeepRun0326AM.xlsx')
        A_KeepRun0326AM =A_KeepRun0326AM.drop_duplicates(subset = None, keep = 'last')
        A_KeepRun0326AM.rename(columns={A_KeepRun0326AM.columns[0]: "Student_ID" }, inplace=True)
        A_KeepRun0326AM.rename(columns={A_KeepRun0326AM.columns[1]:"KeepRun0326AM_In"}, inplace =True)
        A_KeepRun0326AM.rename(columns={A_KeepRun0326AM.columns[2]:"KeepRun0326AM_Out"}, inplace =True)
        
        A_KeepRun0326AM['KeepRun0326AM']=A_KeepRun0326AM['KeepRun0326AM_Out']-A_KeepRun0326AM['KeepRun0326AM_In']
        A_KeepRun0326AM['KeepRun0326AM']=A_KeepRun0326AM['KeepRun0326AM'] / np.timedelta64(60,'s')
        
        '''
        A_KeepRun0326PM
        '''
        A_KeepRun0326PM = pd.read_excel('D:\Mass\Sports_data\Active\Runing\A_KeepRun0326PM.xlsx')
        A_KeepRun0326PM =A_KeepRun0326PM.drop_duplicates(subset = None, keep = 'last')
        A_KeepRun0326PM.rename(columns={A_KeepRun0326PM.columns[0]: "Student_ID" }, inplace=True)
        A_KeepRun0326PM.rename(columns={A_KeepRun0326PM.columns[1]:"KeepRun0326PM_In"}, inplace =True)
        A_KeepRun0326PM.rename(columns={A_KeepRun0326PM.columns[2]:"KeepRun0326PM_Out"}, inplace =True)
        
        A_KeepRun0326PM['KeepRun0326PM']=A_KeepRun0326PM['KeepRun0326PM_Out']-A_KeepRun0326PM['KeepRun0326PM_In']
        A_KeepRun0326PM['KeepRun0326PM']=A_KeepRun0326PM['KeepRun0326PM'] / np.timedelta64(60,'s')
        
        '''
        A_KeepRun0327
        '''
        A_KeepRun0327 = pd.read_excel('D:\Mass\Sports_data\Active\Runing\A_KeepRun0327.xlsx')
        A_KeepRun0327 =A_KeepRun0327.drop_duplicates(subset = None, keep = 'last')
        A_KeepRun0327.rename(columns={A_KeepRun0327.columns[0]: "Student_ID" }, inplace=True)
        A_KeepRun0327.rename(columns={A_KeepRun0327.columns[1]:"KeepRun0327_In"}, inplace =True)
        A_KeepRun0327.rename(columns={A_KeepRun0327.columns[2]:"KeepRun0327_Out"}, inplace =True)
        
        A_KeepRun0327['KeepRun0327']=A_KeepRun0327['KeepRun0327_Out']-A_KeepRun0327['KeepRun0327_In']
        A_KeepRun0327['KeepRun0327']=A_KeepRun0327['KeepRun0327'] / np.timedelta64(60,'s')
        
        '''
        A_KeepRun0328
        '''
        A_KeepRun0328 = pd.read_excel('D:\Mass\Sports_data\Active\Runing\A_KeepRun0328.xlsx')
        A_KeepRun0328 =A_KeepRun0328.drop_duplicates(subset = None, keep = 'last')
        A_KeepRun0328.rename(columns={A_KeepRun0328.columns[0]: "Student_ID" }, inplace=True)
        A_KeepRun0328.rename(columns={A_KeepRun0328.columns[1]:"KeepRun0328_In"}, inplace =True)
        A_KeepRun0328.rename(columns={A_KeepRun0328.columns[2]:"KeepRun0328_Out"}, inplace =True)
        
        A_KeepRun0328['KeepRun0328']=A_KeepRun0328['KeepRun0328_Out']-A_KeepRun0328['KeepRun0328_In']
        A_KeepRun0328['KeepRun0328']=A_KeepRun0328['KeepRun0328'] / np.timedelta64(60,'s')
        
        '''
        A_NightRun20190228
        '''
        A_NightRun0228 =pd.read_excel('D:\Mass\Sports_data\Active\Runing\A_NightRun20190228.xlsx')
        A_NightRun0228 =A_NightRun0228.drop_duplicates(subset = None, keep = 'last')
        A_NightRun0228.rename(columns={A_NightRun0228.columns[0]:"Student_ID"}, inplace=True)
        A_NightRun0228.rename(columns={A_NightRun0228.columns[1]:"NightRun_In0228"}, inplace=True )
        A_NightRun0228.rename(columns={A_NightRun0228.columns[2]:"NightRun_Out0228"},inplace=True)
        
        A_NightRun0228['NightRun0228']=A_NightRun0228['NightRun_Out0228']-A_NightRun0228['NightRun_In0228']
        A_NightRun0228['NightRun0228']=A_NightRun0228['NightRun0228'] / np.timedelta64(60,'s')
        
        
        
        '''
        A_NightRun20190404
        '''
        A_NightRun0404 =pd.read_excel('D:\Mass\Sports_data\Active\Runing\A_NightRun20190404.xlsx')
        A_NightRun0404 =A_NightRun0404.drop_duplicates(subset = None, keep = 'last')
        A_NightRun0404.rename(columns={A_NightRun0404.columns[0]:"Student_ID"}, inplace=True)
        A_NightRun0404.rename(columns={A_NightRun0404.columns[1]:"NightRun_In0404"}, inplace=True )
        A_NightRun0404.rename(columns={A_NightRun0404.columns[2]:"NightRun_Out0404"},inplace=True)
        
        A_NightRun0404['NightRun0404']=A_NightRun0404['NightRun_Out0404']-A_NightRun0404['NightRun_In0404']
        A_NightRun0404['NightRun0404']=A_NightRun0404['NightRun0404'] / np.timedelta64(60,'s')
        
        
        '''
        A_NightRun20190416
        '''
        A_NightRun0416 =pd.read_excel('D:\Mass\Sports_data\Active\Runing\A_NightRun20190416.xlsx')
        A_NightRun0416 =A_NightRun0416.drop_duplicates(subset = None, keep = 'last')
        A_NightRun0416.rename(columns={A_NightRun0416.columns[0]:"Student_ID"}, inplace=True)
        A_NightRun0416.rename(columns={A_NightRun0416.columns[1]:"NightRun_In0416"}, inplace=True )
        A_NightRun0416.rename(columns={A_NightRun0416.columns[2]:"NightRun_Out0416"},inplace=True)
        
        A_NightRun0416['NightRun0416']=A_NightRun0416['NightRun_Out0416']-A_NightRun0416['NightRun_In0416']
        A_NightRun0416['NightRun0416']=A_NightRun0416['NightRun0416'] / np.timedelta64(60,'s')
        
        
        '''
        A_NightWalk
        '''
        A_NightWalk =pd.read_excel('D:\Mass\Sports_data\Active\Runing\A_NightWalk.xlsx')
        A_NightWalk =A_NightWalk.drop_duplicates(subset = None, keep = 'last')
        A_NightWalk.rename(columns={A_NightWalk.columns[0]:"Student_ID"}, inplace=True)
        A_NightWalk.rename(columns={A_NightWalk.columns[1]:"NightWalk"}, inplace=True )
        A_NightWalk['NightWalk']=A_NightWalk['NightWalk']  * 60
        
        '''
        A_Letexercise_session1
        '''
        A_Letexercise_session1 =pd.read_excel('D:\Mass\Sports_data\Active\Runing\A_Letexercise_session1.xlsx')
        A_Letexercise_session1 =A_Letexercise_session1.drop_duplicates(subset = None, keep = 'last')
        A_Letexercise_session1.rename(columns={A_Letexercise_session1.columns[0]:"Student_ID"}, inplace=True)
        A_Letexercise_session1.rename(columns={A_Letexercise_session1.columns[1]:"Letexercise1_In"}, inplace=True )
        A_Letexercise_session1.rename(columns={A_Letexercise_session1.columns[2]:"Letexercise1_Out"},inplace=True)
        
        A_Letexercise_session1['Letexercise_session1']=A_Letexercise_session1['Letexercise1_Out']-A_Letexercise_session1['Letexercise1_In']
        A_Letexercise_session1['Letexercise_session1']=A_Letexercise_session1['Letexercise_session1'] / np.timedelta64(60,'s')
        
        '''
        A_Letexercise_session2
        '''
        A_Letexercise_session2 =pd.read_excel('D:\Mass\Sports_data\Active\Runing\A_Letexercise_session2.xlsx')
        A_Letexercise_session2 =A_Letexercise_session2.drop_duplicates(subset = None, keep = 'last')
        A_Letexercise_session2.rename(columns={A_Letexercise_session2.columns[0]:"Student_ID"}, inplace=True)
        A_Letexercise_session2.rename(columns={A_Letexercise_session2.columns[1]:"Letexercise2_In"}, inplace=True )
        A_Letexercise_session2.rename(columns={A_Letexercise_session2.columns[2]:"Letexercise2_Out"},inplace=True)
        
        A_Letexercise_session2['Letexercise_session2']=A_Letexercise_session2['Letexercise2_Out']-A_Letexercise_session2['Letexercise2_In']
        A_Letexercise_session2['Letexercise_session2']=A_Letexercise_session2['Letexercise_session2'] / np.timedelta64(60,'s')
        
        '''
        A_Letexercise_session3
        '''
        A_Letexercise_session3 =pd.read_excel('D:\Mass\Sports_data\Active\Runing\A_Letexercise_session3.xlsx')
        A_Letexercise_session3 =A_Letexercise_session3.drop_duplicates(subset = None, keep = 'last')
        A_Letexercise_session3.rename(columns={A_Letexercise_session3.columns[0]:"Student_ID"}, inplace=True)
        A_Letexercise_session3.rename(columns={A_Letexercise_session3.columns[1]:"Letexercise3_In"}, inplace=True )
        A_Letexercise_session3.rename(columns={A_Letexercise_session3.columns[2]:"Letexercise3_Out"},inplace=True)
        
        A_Letexercise_session3['Letexercise_session3']=A_Letexercise_session3['Letexercise3_Out']-A_Letexercise_session3['Letexercise3_In']
        A_Letexercise_session3['Letexercise_session3']=A_Letexercise_session3['Letexercise_session3'] / np.timedelta64(60,'s')
        
        '''
        A_Letexercise_session4
        '''
        A_Letexercise_session4 =pd.read_excel('D:\Mass\Sports_data\Active\Runing\A_Letexercise_session4.xlsx')
        A_Letexercise_session4 =A_Letexercise_session4.drop_duplicates(subset = None, keep = 'last')
        A_Letexercise_session4.rename(columns={A_Letexercise_session4.columns[0]:"Student_ID"}, inplace=True)
        A_Letexercise_session4.rename(columns={A_Letexercise_session4.columns[1]:"Letexercise4_In"}, inplace=True )
        A_Letexercise_session4.rename(columns={A_Letexercise_session4.columns[2]:"Letexercise4_Out"},inplace=True)
        
        A_Letexercise_session4['Letexercise_session4']=A_Letexercise_session4['Letexercise4_Out']-A_Letexercise_session4['Letexercise4_In']
        A_Letexercise_session4['Letexercise_session4']=A_Letexercise_session4['Letexercise_session4'] / np.timedelta64(60,'s')
        
        '''
        A_Letexercise_session5
        '''
        A_Letexercise_session5 =pd.read_excel('D:\Mass\Sports_data\Active\Runing\A_Letexercise_session5.xlsx')
        A_Letexercise_session5 =A_Letexercise_session5.drop_duplicates(subset = None, keep = 'last')
        A_Letexercise_session5.rename(columns={A_Letexercise_session5.columns[0]:"Student_ID"}, inplace=True)
        A_Letexercise_session5.rename(columns={A_Letexercise_session5.columns[1]:"Letexercise5_In"}, inplace=True )
        A_Letexercise_session5.rename(columns={A_Letexercise_session5.columns[2]:"Letexercise5_Out"},inplace=True)
        
        A_Letexercise_session5['Letexercise_session5']=A_Letexercise_session5['Letexercise5_Out']-A_Letexercise_session5['Letexercise5_In']
        A_Letexercise_session5['Letexercise_session5']=A_Letexercise_session5['Letexercise_session5'] / np.timedelta64(60,'s')
        
        '''
        A_Letexercise_session6
        '''
        A_Letexercise_session6 =pd.read_excel('D:\Mass\Sports_data\Active\Runing\A_Letexercise_session6.xlsx')
        A_Letexercise_session6 =A_Letexercise_session6.drop_duplicates(subset = None, keep = 'last')
        A_Letexercise_session6.rename(columns={A_Letexercise_session6.columns[0]:"Student_ID"}, inplace=True)
        A_Letexercise_session6.rename(columns={A_Letexercise_session6.columns[1]:"Letexercise6_In"}, inplace=True )
        A_Letexercise_session6.rename(columns={A_Letexercise_session6.columns[2]:"Letexercise6_Out"},inplace=True)
        
        A_Letexercise_session6['Letexercise_session6']=A_Letexercise_session6['Letexercise6_Out']-A_Letexercise_session6['Letexercise6_In']
        A_Letexercise_session6['Letexercise_session6']=A_Letexercise_session6['Letexercise_session6'] / np.timedelta64(60,'s')
        
        '''
        A_Letexercise_session7
        '''
        A_Letexercise_session7 =pd.read_excel('D:\Mass\Sports_data\Active\Runing\A_Letexercise_session7.xlsx')
        A_Letexercise_session7 =A_Letexercise_session7.drop_duplicates(subset = None, keep = 'last')
        A_Letexercise_session7.rename(columns={A_Letexercise_session7.columns[0]:"Student_ID"}, inplace=True)
        A_Letexercise_session7.rename(columns={A_Letexercise_session7.columns[1]:"Letexercise7_In"}, inplace=True )
        A_Letexercise_session7.rename(columns={A_Letexercise_session7.columns[2]:"Letexercise7_Out"},inplace=True)
        
        A_Letexercise_session7['Letexercise_session7']=A_Letexercise_session7['Letexercise7_Out']-A_Letexercise_session7['Letexercise7_In']
        A_Letexercise_session7['Letexercise_session7']=A_Letexercise_session7['Letexercise_session7'] / np.timedelta64(60,'s')
        
        '''
        A_Letexercise_session8
        '''
        A_Letexercise_session8 =pd.read_excel('D:\Mass\Sports_data\Active\Runing\A_Letexercise_session8.xlsx')
        A_Letexercise_session8 =A_Letexercise_session8.drop_duplicates(subset = None, keep = 'last')
        A_Letexercise_session8.rename(columns={A_Letexercise_session8.columns[0]:"Student_ID"}, inplace=True)
        A_Letexercise_session8.rename(columns={A_Letexercise_session8.columns[1]:"Letexercise8_In"}, inplace=True )
        A_Letexercise_session8.rename(columns={A_Letexercise_session8.columns[2]:"Letexercise8_Out"},inplace=True)
        
        A_Letexercise_session8['Letexercise_session8']=A_Letexercise_session8['Letexercise8_Out']-A_Letexercise_session8['Letexercise8_In']
        A_Letexercise_session8['Letexercise_session8']=A_Letexercise_session8['Letexercise_session8'] / np.timedelta64(60,'s')
        
        '''
        A_Letexercise_session9
        '''
        A_Letexercise_session9 =pd.read_excel('D:\Mass\Sports_data\Active\Runing\A_Letexercise_session9.xlsx')
        A_Letexercise_session9 =A_Letexercise_session9.drop_duplicates(subset = None, keep = 'last')
        A_Letexercise_session9.rename(columns={A_Letexercise_session9.columns[0]:"Student_ID"}, inplace=True)
        A_Letexercise_session9.rename(columns={A_Letexercise_session9.columns[1]:"Letexercise9_In"}, inplace=True )
        A_Letexercise_session9.rename(columns={A_Letexercise_session9.columns[2]:"Letexercise9_Out"},inplace=True)
        
        A_Letexercise_session9['Letexercise_session9']=A_Letexercise_session9['Letexercise9_Out']-A_Letexercise_session9['Letexercise9_In']
        A_Letexercise_session9['Letexercise_session9']=A_Letexercise_session9['Letexercise_session9'] / np.timedelta64(60,'s')
        
        '''
        A_Letexercise_session10
        '''
        A_Letexercise_session10 =pd.read_excel('D:\Mass\Sports_data\Active\Runing\A_Letexercise_session10.xlsx')
        A_Letexercise_session10 =A_Letexercise_session10.drop_duplicates(subset = None, keep = 'last')
        A_Letexercise_session10.rename(columns={A_Letexercise_session10.columns[0]:"Student_ID"}, inplace=True)
        A_Letexercise_session10.rename(columns={A_Letexercise_session10.columns[1]:"Letexercise10_In"}, inplace=True )
        A_Letexercise_session10.rename(columns={A_Letexercise_session10.columns[2]:"Letexercise10_Out"},inplace=True)
        
        A_Letexercise_session10['Letexercise_session10']=A_Letexercise_session10['Letexercise10_Out']-A_Letexercise_session10['Letexercise10_In']
        A_Letexercise_session10['Letexercise_session10']=A_Letexercise_session10['Letexercise_session10'] / np.timedelta64(60,'s')
        
        '''
        A_Letexercise_session11
        '''
        A_Letexercise_session11 =pd.read_excel('D:\Mass\Sports_data\Active\Runing\A_Letexercise_session11.xlsx')
        A_Letexercise_session11 =A_Letexercise_session11.drop_duplicates(subset = None, keep = 'last')
        A_Letexercise_session11.rename(columns={A_Letexercise_session11.columns[0]:"Student_ID"}, inplace=True)
        A_Letexercise_session11.rename(columns={A_Letexercise_session11.columns[1]:"Letexercise11_In"}, inplace=True )
        A_Letexercise_session11.rename(columns={A_Letexercise_session11.columns[2]:"Letexercise11_Out"},inplace=True)
        
        A_Letexercise_session11['Letexercise_session11']=A_Letexercise_session11['Letexercise11_Out']-A_Letexercise_session11['Letexercise11_In']
        A_Letexercise_session11['Letexercise_session11']=A_Letexercise_session11['Letexercise_session11'] / np.timedelta64(60,'s')
        
        '''
        A_Letexercise_session12
        '''
        A_Letexercise_session12 =pd.read_excel('D:\Mass\Sports_data\Active\Runing\A_Letexercise_session12.xlsx')
        A_Letexercise_session12 =A_Letexercise_session12.drop_duplicates(subset = None, keep = 'last')
        A_Letexercise_session12.rename(columns={A_Letexercise_session12.columns[0]:"Student_ID"}, inplace=True)
        A_Letexercise_session12.rename(columns={A_Letexercise_session12.columns[1]:"Letexercise12_In"}, inplace=True )
        A_Letexercise_session12.rename(columns={A_Letexercise_session12.columns[2]:"Letexercise12_Out"},inplace=True)
        
        A_Letexercise_session12['Letexercise_session12']=A_Letexercise_session12['Letexercise12_Out']-A_Letexercise_session12['Letexercise12_In']
        A_Letexercise_session12['Letexercise_session12']=A_Letexercise_session12['Letexercise_session12'] / np.timedelta64(60,'s')
        
        '''
        A_Letexercise_session13
        '''
        A_Letexercise_session13 =pd.read_excel('D:\Mass\Sports_data\Active\Runing\A_Letexercise_session13.xlsx')
        A_Letexercise_session13 =A_Letexercise_session13.drop_duplicates(subset = None, keep = 'last')
        A_Letexercise_session13.rename(columns={A_Letexercise_session13.columns[0]:"Student_ID"}, inplace=True)
        A_Letexercise_session13.rename(columns={A_Letexercise_session13.columns[1]:"Letexercise13_In"}, inplace=True )
        A_Letexercise_session13.rename(columns={A_Letexercise_session13.columns[2]:"Letexercise13_Out"},inplace=True)
        
        A_Letexercise_session13['Letexercise_session13']=A_Letexercise_session13['Letexercise13_Out']-A_Letexercise_session13['Letexercise13_In']
        A_Letexercise_session13['Letexercise_session13']=A_Letexercise_session13['Letexercise_session13'] / np.timedelta64(60,'s')
        
        '''
        A_Letexercise_session14
        '''
        A_Letexercise_session14 =pd.read_excel('D:\Mass\Sports_data\Active\Runing\A_Letexercise_session14.xlsx')
        A_Letexercise_session14 =A_Letexercise_session14.drop_duplicates(subset = None, keep = 'last')
        A_Letexercise_session14.rename(columns={A_Letexercise_session14.columns[0]:"Student_ID"}, inplace=True)
        A_Letexercise_session14.rename(columns={A_Letexercise_session14.columns[1]:"Letexercise14_In"}, inplace=True )
        A_Letexercise_session14.rename(columns={A_Letexercise_session14.columns[2]:"Letexercise14_Out"},inplace=True)
        
        A_Letexercise_session14['Letexercise_session14']=A_Letexercise_session14['Letexercise14_Out']-A_Letexercise_session14['Letexercise14_In']
        A_Letexercise_session14['Letexercise_session14']=A_Letexercise_session14['Letexercise_session14'] / np.timedelta64(60,'s')
        
        '''
        A_Letexercise_session15
        '''
        A_Letexercise_session15 =pd.read_excel('D:\Mass\Sports_data\Active\Runing\A_Letexercise_session15.xlsx')
        A_Letexercise_session15 =A_Letexercise_session15.drop_duplicates(subset = None, keep = 'last')
        A_Letexercise_session15.rename(columns={A_Letexercise_session15.columns[0]:"Student_ID"}, inplace=True)
        A_Letexercise_session15.rename(columns={A_Letexercise_session15.columns[1]:"Letexercise_session15"}, inplace=True )
        A_Letexercise_session15['Letexercise_session15']=A_Letexercise_session15['Letexercise_session15']  * 60
        
        '''
        A_Letexercise_session16
        '''
        A_Letexercise_session16 =pd.read_excel('D:\Mass\Sports_data\Active\Runing\A_Letexercise_session16.xlsx')
        A_Letexercise_session16 =A_Letexercise_session16.drop_duplicates(subset = None, keep = 'last')
        A_Letexercise_session16.rename(columns={A_Letexercise_session16.columns[0]:"Student_ID"}, inplace=True)
        A_Letexercise_session16.rename(columns={A_Letexercise_session16.columns[1]:"Letexercise_session16"}, inplace=True )
        A_Letexercise_session16['Letexercise_session16']=A_Letexercise_session16['Letexercise_session16']  * 60
        
        '''
        A_Letexercise_session17
        '''
        A_Letexercise_session17 =pd.read_excel('D:\Mass\Sports_data\Active\Runing\A_Letexercise_session17.xlsx')
        A_Letexercise_session17 =A_Letexercise_session17.drop_duplicates(subset = None, keep = 'last')
        A_Letexercise_session17.rename(columns={A_Letexercise_session17.columns[0]:"Student_ID"}, inplace=True)
        A_Letexercise_session17.rename(columns={A_Letexercise_session17.columns[1]:"Letexercise_session17"}, inplace=True )
        A_Letexercise_session17['Letexercise_session17']=A_Letexercise_session17['Letexercise_session17']  * 60
        
        '''
        A_Letexercise_session18
        '''
        A_Letexercise_session18 =pd.read_excel('D:\Mass\Sports_data\Active\Runing\A_Letexercise_session18.xlsx')
        A_Letexercise_session18 =A_Letexercise_session18.drop_duplicates(subset = None, keep = 'last')
        A_Letexercise_session18.rename(columns={A_Letexercise_session18.columns[0]:"Student_ID"}, inplace=True)
        A_Letexercise_session18.rename(columns={A_Letexercise_session18.columns[1]:"Letexercise_session18"}, inplace=True )
        A_Letexercise_session18['Letexercise_session18']=A_Letexercise_session18['Letexercise_session18']  * 60
        
        '''
        A_Letexercise_session19
        '''
        A_Letexercise_session19 =pd.read_excel('D:\Mass\Sports_data\Active\Runing\A_Letexercise_session19.xlsx')
        A_Letexercise_session19 =A_Letexercise_session19.drop_duplicates(subset = None, keep = 'last')
        A_Letexercise_session19.rename(columns={A_Letexercise_session19.columns[0]:"Student_ID"}, inplace=True)
        A_Letexercise_session19.rename(columns={A_Letexercise_session19.columns[1]:"Letexercise_session19"}, inplace=True )
        A_Letexercise_session19['Letexercise_session19']=A_Letexercise_session19['Letexercise_session19']  * 60
        
        '''
        A_Letexercise_session20
        '''
        A_Letexercise_session20 =pd.read_excel('D:\Mass\Sports_data\Active\Runing\A_Letexercise_session20.xlsx')
        A_Letexercise_session20 =A_Letexercise_session20.drop_duplicates(subset = None, keep = 'last')
        A_Letexercise_session20.rename(columns={A_Letexercise_session20.columns[0]:"Student_ID"}, inplace=True)
        A_Letexercise_session20.rename(columns={A_Letexercise_session20.columns[1]:"Letexercise_session20"}, inplace=True )
        A_Letexercise_session20['Letexercise_session20']=A_Letexercise_session20['Letexercise_session20']  * 60
        
        '''
        A_Letexercise_session21
        '''
        A_Letexercise_session21 =pd.read_excel('D:\Mass\Sports_data\Active\Runing\A_Letexercise_session21.xlsx')
        A_Letexercise_session21 =A_Letexercise_session21.drop_duplicates(subset = None, keep = 'last')
        A_Letexercise_session21.rename(columns={A_Letexercise_session21.columns[0]:"Student_ID"}, inplace=True)
        A_Letexercise_session21.rename(columns={A_Letexercise_session21.columns[1]:"Letexercise_session21"}, inplace=True )
        A_Letexercise_session21['Letexercise_session21']=A_Letexercise_session21['Letexercise_session21']  * 60
        
        '''
        A_Letexercise_session22
        '''
        A_Letexercise_session22 =pd.read_excel('D:\Mass\Sports_data\Active\Runing\A_Letexercise_session22.xlsx')
        A_Letexercise_session22 =A_Letexercise_session22.drop_duplicates(subset = None, keep = 'last')
        A_Letexercise_session22.rename(columns={A_Letexercise_session22.columns[0]:"Student_ID"}, inplace=True)
        A_Letexercise_session22.rename(columns={A_Letexercise_session22.columns[1]:"Letexercise_session22"}, inplace=True )
        A_Letexercise_session22['Letexercise_session22']=A_Letexercise_session22['Letexercise_session22']  * 60
        
        '''
        A_Letexercise_session23
        '''
        A_Letexercise_session23 =pd.read_excel('D:\Mass\Sports_data\Active\Runing\A_Letexercise_session23.xlsx')
        A_Letexercise_session23 =A_Letexercise_session23.drop_duplicates(subset = None, keep = 'last')
        A_Letexercise_session23.rename(columns={A_Letexercise_session23.columns[0]:"Student_ID"}, inplace=True)
        A_Letexercise_session23.rename(columns={A_Letexercise_session23.columns[1]:"Letexercise_session23"}, inplace=True )
        A_Letexercise_session23['Letexercise_session23']=A_Letexercise_session23['Letexercise_session23']  * 60
        
        '''
        A_Letexercise_session24
        '''
        A_Letexercise_session24 =pd.read_excel('D:\Mass\Sports_data\Active\Runing\A_Letexercise_session24.xlsx')
        A_Letexercise_session24 =A_Letexercise_session24.drop_duplicates(subset = None, keep = 'last')
        A_Letexercise_session24.rename(columns={A_Letexercise_session24.columns[0]:"Student_ID"}, inplace=True)
        A_Letexercise_session24.rename(columns={A_Letexercise_session24.columns[1]:"Letexercise_session24"}, inplace=True )
        A_Letexercise_session24['Letexercise_session24']=A_Letexercise_session24['Letexercise_session24']  * 60
        
        '''
        A_DiaboloCourse0228
        '''
        A_DiaboloCourse0228 = pd.read_excel('D:\Mass\Sports_data\Active\Courses\A_DiaboloCourse0228.xlsx')
        A_DiaboloCourse0228 =A_DiaboloCourse0228.drop_duplicates(subset = None, keep = 'last')
        A_DiaboloCourse0228.rename(columns={A_DiaboloCourse0228.columns[0]: "Student_ID" }, inplace=True)
        A_DiaboloCourse0228.rename(columns={A_DiaboloCourse0228.columns[1]:"DiaboloCourse0228_In"}, inplace =True)
        A_DiaboloCourse0228.rename(columns={A_DiaboloCourse0228.columns[2]:"DiaboloCourse0228_Out"}, inplace =True)
        
        A_DiaboloCourse0228['DiaboloCourse0228']=A_DiaboloCourse0228['DiaboloCourse0228_Out']-A_DiaboloCourse0228['DiaboloCourse0228_In']
        A_DiaboloCourse0228['DiaboloCourse0228']=A_DiaboloCourse0228['DiaboloCourse0228'] / np.timedelta64(60,'s')
        
        
        '''
        A_DiaboloCourse0307
        '''
        A_DiaboloCourse0307 = pd.read_excel('D:\Mass\Sports_data\Active\Courses\A_DiaboloCourse0307.xlsx')
        A_DiaboloCourse0307 =A_DiaboloCourse0307.drop_duplicates(subset = None, keep = 'last')
        A_DiaboloCourse0307.rename(columns={A_DiaboloCourse0307.columns[0]: "Student_ID" }, inplace=True)
        A_DiaboloCourse0307.rename(columns={A_DiaboloCourse0307.columns[1]:"DiaboloCourse0307_In"}, inplace =True)
        A_DiaboloCourse0307.rename(columns={A_DiaboloCourse0307.columns[2]:"DiaboloCourse0307_Out"}, inplace =True)
        
        A_DiaboloCourse0307['DiaboloCourse0307']=A_DiaboloCourse0307['DiaboloCourse0307_Out']-A_DiaboloCourse0307['DiaboloCourse0307_In']
        A_DiaboloCourse0307['DiaboloCourse0307']=A_DiaboloCourse0307['DiaboloCourse0307'] / np.timedelta64(60,'s')
        
        '''
        A_DiaboloCourse0314
        '''
        A_DiaboloCourse0314 = pd.read_excel('D:\Mass\Sports_data\Active\Courses\A_DiaboloCourse0314.xlsx')
        A_DiaboloCourse0314 =A_DiaboloCourse0314.drop_duplicates(subset = None, keep = 'last')
        A_DiaboloCourse0314.rename(columns={A_DiaboloCourse0314.columns[0]: "Student_ID" }, inplace=True)
        A_DiaboloCourse0314.rename(columns={A_DiaboloCourse0314.columns[1]:"DiaboloCourse0314_In"}, inplace =True)
        A_DiaboloCourse0314.rename(columns={A_DiaboloCourse0314.columns[2]:"DiaboloCourse0314_Out"}, inplace =True)
        
        A_DiaboloCourse0314['DiaboloCourse0314']=A_DiaboloCourse0314['DiaboloCourse0314_Out']-A_DiaboloCourse0314['DiaboloCourse0314_In']
        A_DiaboloCourse0314['DiaboloCourse0314']=A_DiaboloCourse0314['DiaboloCourse0314'] / np.timedelta64(60,'s')
        
        '''
        A_DiaboloCourse0321
        '''
        A_DiaboloCourse0321 = pd.read_excel('D:\Mass\Sports_data\Active\Courses\A_DiaboloCourse0321.xlsx')
        A_DiaboloCourse0321 =A_DiaboloCourse0321.drop_duplicates(subset = None, keep = 'last')
        A_DiaboloCourse0321.rename(columns={A_DiaboloCourse0321.columns[0]: "Student_ID" }, inplace=True)
        A_DiaboloCourse0321.rename(columns={A_DiaboloCourse0321.columns[1]:"DiaboloCourse0321_In"}, inplace =True)
        A_DiaboloCourse0321.rename(columns={A_DiaboloCourse0321.columns[2]:"DiaboloCourse0321_Out"}, inplace =True)
        
        A_DiaboloCourse0321['DiaboloCourse0321']=A_DiaboloCourse0321['DiaboloCourse0321_Out']-A_DiaboloCourse0321['DiaboloCourse0321_In']
        A_DiaboloCourse0321['DiaboloCourse0321']=A_DiaboloCourse0321['DiaboloCourse0321'] / np.timedelta64(60,'s')
        
        '''
        A_DiaboloCourse0328
        '''
        A_DiaboloCourse0328 = pd.read_excel('D:\Mass\Sports_data\Active\Courses\A_DiaboloCourse0328.xlsx')
        A_DiaboloCourse0328 =A_DiaboloCourse0328.drop_duplicates(subset = None, keep = 'last')
        A_DiaboloCourse0328.rename(columns={A_DiaboloCourse0328.columns[0]: "Student_ID" }, inplace=True)
        A_DiaboloCourse0328.rename(columns={A_DiaboloCourse0328.columns[1]:"DiaboloCourse0328_In"}, inplace =True)
        A_DiaboloCourse0328.rename(columns={A_DiaboloCourse0328.columns[2]:"DiaboloCourse0328_Out"}, inplace =True)
        
        A_DiaboloCourse0328['DiaboloCourse0328']=A_DiaboloCourse0328['DiaboloCourse0328_Out']-A_DiaboloCourse0328['DiaboloCourse0328_In']
        A_DiaboloCourse0328['DiaboloCourse0328']=A_DiaboloCourse0328['DiaboloCourse0328'] / np.timedelta64(60,'s')
        
        '''
        A_Tennis0320
        '''
        A_Tennis0320 = pd.read_excel('D:\Mass\Sports_data\Active\Courses\A_Tennis0320.xlsx')
        A_Tennis0320 =A_Tennis0320.drop_duplicates(subset = None, keep = 'last')
        A_Tennis0320.rename(columns={A_Tennis0320.columns[0]: "Student_ID" }, inplace=True)
        A_Tennis0320.rename(columns={A_Tennis0320.columns[1]:"Tennis0320_In"}, inplace =True)
        A_Tennis0320.rename(columns={A_Tennis0320.columns[2]:"Tennis0320_Out"}, inplace =True)
        
        A_Tennis0320['Tennis0320']=A_Tennis0320['Tennis0320_Out']-A_Tennis0320['Tennis0320_In']
        A_Tennis0320['Tennis0320']=A_Tennis0320['Tennis0320'] / np.timedelta64(60,'s')
        
        '''
        A_Tennis0327
        '''
        A_Tennis0327 = pd.read_excel('D:\Mass\Sports_data\Active\Courses\A_Tennis0327.xlsx')
        A_Tennis0327 =A_Tennis0327.drop_duplicates(subset = None, keep = 'last')
        A_Tennis0327.rename(columns={A_Tennis0327.columns[0]: "Student_ID" }, inplace=True)
        A_Tennis0327.rename(columns={A_Tennis0327.columns[1]:"Tennis0327_In"}, inplace =True)
        A_Tennis0327.rename(columns={A_Tennis0327.columns[2]:"Tennis0327_Out"}, inplace =True)
        
        A_Tennis0327['Tennis0327']=A_Tennis0327['Tennis0327_Out']-A_Tennis0327['Tennis0327_In']
        A_Tennis0327['Tennis0327']=A_Tennis0327['Tennis0327'] / np.timedelta64(60,'s')
        
        '''
        A_TableTennis04
        '''
        A_TableTennis04 = pd.read_excel('D:\Mass\Sports_data\Active\Courses\A_TableTennis04.xlsx')
        A_TableTennis04 =A_TableTennis04.drop_duplicates(subset = None, keep = 'last')
        A_TableTennis04.rename(columns={A_TableTennis04.columns[0]: "Student_ID" }, inplace=True)
        A_TableTennis04.rename(columns={A_TableTennis04.columns[1]:"TableTennis04"}, inplace =True)
        A_TableTennis04['TableTennis04']=A_TableTennis04['TableTennis04'] * 60
        
        '''
        A_TennisInterestCourse
        '''
        A_TennisInterestCourse = pd.read_excel('D:\Mass\Sports_data\Active\Courses\A_TennisInterestCourse.xlsx')
        A_TennisInterestCourse =A_TennisInterestCourse.drop_duplicates(subset = None, keep = 'last')
        A_TennisInterestCourse.rename(columns={A_TennisInterestCourse.columns[0]: "Student_ID" }, inplace=True)
        A_TennisInterestCourse.rename(columns={A_TennisInterestCourse.columns[1]:"TennisInterestCourse"}, inplace =True)
        A_TennisInterestCourse['TennisInterestCourse']=A_TennisInterestCourse['TennisInterestCourse'] * 60
        
        
        '''
        A_WomenSoccor
        '''
        A_WomenSoccor = pd.read_excel('D:\Mass\Sports_data\Active\Courses\A_WomenSoccor.xlsx')
        A_WomenSoccor =A_WomenSoccor.drop_duplicates(subset = None, keep = 'last')
        A_WomenSoccor.rename(columns={A_WomenSoccor.columns[0]: "Student_ID" }, inplace=True)
        A_WomenSoccor.rename(columns={A_WomenSoccor.columns[1]:"WomenSoccor_In"}, inplace =True)
        A_WomenSoccor.rename(columns={A_WomenSoccor.columns[2]:"WomenSoccor_Out"}, inplace =True)
        
        A_WomenSoccor['WomenSoccor']=A_WomenSoccor['WomenSoccor_Out']-A_WomenSoccor['WomenSoccor_In']
        A_WomenSoccor['WomenSoccor']=A_WomenSoccor['WomenSoccor'] / np.timedelta64(60,'s')
        
         
        '''
        A_UMSUTrack_Field
        '''
        A_UMSUTrack_Field = pd.read_excel('D:\Mass\Sports_data\Active\Courses\A_UMSUTrack_Field.xlsx')
        A_UMSUTrack_Field =A_UMSUTrack_Field.drop_duplicates(subset = None, keep = 'last')
        A_UMSUTrack_Field.rename(columns={A_UMSUTrack_Field.columns[0]: "Student_ID" }, inplace=True)
        A_UMSUTrack_Field.rename(columns={A_UMSUTrack_Field.columns[1]:"UMSUTrack_Field_In"}, inplace =True)
        A_UMSUTrack_Field.rename(columns={A_UMSUTrack_Field.columns[2]:"UMSUTrack_Field_Out"}, inplace =True)
        
        A_UMSUTrack_Field['UMSUTrack_Field']=A_UMSUTrack_Field['UMSUTrack_Field_Out']-A_UMSUTrack_Field['UMSUTrack_Field_In']
        A_UMSUTrack_Field['UMSUTrack_Field']=A_UMSUTrack_Field['UMSUTrack_Field'] / np.timedelta64(60,'s')
        
        '''
        A_SwimTime
        '''
        A_SwimTime = pd.read_excel('D:\Mass\Sports_data\Active\Courses\A_SwimTime.xlsx')
        A_SwimTime =A_SwimTime.drop_duplicates(subset = None, keep = 'last')
        A_SwimTime.rename(columns={A_SwimTime.columns[0]: "Student_ID" }, inplace=True)
        A_SwimTime.rename(columns={A_SwimTime.columns[1]:"SwimTime"}, inplace =True)
        A_SwimTime['SwimTime']=A_SwimTime['SwimTime'] * 60
        
        '''
        A_SwimTime04
        '''
        A_SwimTime04 = pd.read_excel('D:\Mass\Sports_data\Active\Courses\A_SwimTime04.xlsx')
        A_SwimTime04 =A_SwimTime04.drop_duplicates(subset = None, keep = 'last')
        A_SwimTime04.rename(columns={A_SwimTime04.columns[0]: "Student_ID" }, inplace=True)
        A_SwimTime04.rename(columns={A_SwimTime04.columns[1]:"SwimTime04"}, inplace =True)
        A_SwimTime04['SwimTime04']=A_SwimTime04['SwimTime04'] * 60
        
        '''
        A_Korfball0413
        '''
        A_Korfball0413 = pd.read_excel('D:\Mass\Sports_data\Active\Courses\A_Korfball0413.xlsx')
        A_Korfball0413 =A_Korfball0413.drop_duplicates(subset = None, keep = 'last')
        A_Korfball0413.rename(columns={A_Korfball0413.columns[0]: "Student_ID" }, inplace=True)
        A_Korfball0413.rename(columns={A_Korfball0413.columns[1]:"Korfball0413"}, inplace =True)
        A_Korfball0413['Korfball0413']=A_Korfball0413['Korfball0413'] * 60
        
        
        '''
        A_Korfball0420
        '''
        A_Korfball0420 = pd.read_excel('D:\Mass\Sports_data\Active\Courses\A_Korfball0420.xlsx')
        A_Korfball0420 =A_Korfball0420.drop_duplicates(subset = None, keep = 'last')
        A_Korfball0420.rename(columns={A_Korfball0420.columns[0]: "Student_ID" }, inplace=True)
        A_Korfball0420.rename(columns={A_Korfball0420.columns[1]:"Korfball0420"}, inplace =True)
        A_Korfball0420['Korfball0420']=A_Korfball0420['Korfball0420'] * 60
        
        
        '''
        A_KoreanCulture0309
        '''
        A_KoreanCulture0309 = pd.read_excel('D:\Mass\Sports_data\Active\Courses\A_KoreanCulture0309.xlsx')
        A_KoreanCulture0309 =A_KoreanCulture0309.drop_duplicates(subset = None, keep = 'last')
        A_KoreanCulture0309.rename(columns={A_KoreanCulture0309.columns[0]: "Student_ID" }, inplace=True)
        A_KoreanCulture0309.rename(columns={A_KoreanCulture0309.columns[1]:"KoreanCulture0309"}, inplace =True)
        A_KoreanCulture0309['KoreanCulture0309']=A_KoreanCulture0309['KoreanCulture0309'] * 60
        
        '''
        A_KoreanCulture0316
        '''
        A_KoreanCulture0316 = pd.read_excel('D:\Mass\Sports_data\Active\Courses\A_KoreanCulture0316.xlsx')
        A_KoreanCulture0316 =A_KoreanCulture0316.drop_duplicates(subset = None, keep = 'last')
        A_KoreanCulture0316.rename(columns={A_KoreanCulture0316.columns[0]: "Student_ID" }, inplace=True)
        A_KoreanCulture0316.rename(columns={A_KoreanCulture0316.columns[1]:"KoreanCulture0316"}, inplace =True)
        A_KoreanCulture0316['KoreanCulture0316']=A_KoreanCulture0316['KoreanCulture0316'] * 60
        
        '''
        A_KoreanCulture0323
        '''
        A_KoreanCulture0323 = pd.read_excel('D:\Mass\Sports_data\Active\Courses\A_KoreanCulture0323.xlsx')
        A_KoreanCulture0323 =A_KoreanCulture0323.drop_duplicates(subset = None, keep = 'last')
        A_KoreanCulture0323.rename(columns={A_KoreanCulture0323.columns[0]: "Student_ID" }, inplace=True)
        A_KoreanCulture0323.rename(columns={A_KoreanCulture0323.columns[1]:"KoreanCulture0323"}, inplace =True)
        A_KoreanCulture0323['KoreanCulture0323']=A_KoreanCulture0323['KoreanCulture0323'] * 60
        
        '''
        A_KoreanCulture0330
        '''
        A_KoreanCulture0330 = pd.read_excel('D:\Mass\Sports_data\Active\Courses\A_KoreanCulture0330.xlsx')
        A_KoreanCulture0330 =A_KoreanCulture0330.drop_duplicates(subset = None, keep = 'last')
        A_KoreanCulture0330.rename(columns={A_KoreanCulture0330.columns[0]: "Student_ID" }, inplace=True)
        A_KoreanCulture0330.rename(columns={A_KoreanCulture0330.columns[1]:"KoreanCulture0330"}, inplace =True)
        A_KoreanCulture0330['KoreanCulture0330']=A_KoreanCulture0330['KoreanCulture0330'] * 60
        
        '''
        A_KoreanCulture0403
        '''
        A_KoreanCulture0403 = pd.read_excel('D:\Mass\Sports_data\Active\Courses\A_KoreanCulture0403.xlsx')
        A_KoreanCulture0403 =A_KoreanCulture0403.drop_duplicates(subset = None, keep = 'last')
        A_KoreanCulture0403.rename(columns={A_KoreanCulture0403.columns[0]: "Student_ID" }, inplace=True)
        A_KoreanCulture0403.rename(columns={A_KoreanCulture0403.columns[1]:"KoreanCulture0403"}, inplace =True)
        A_KoreanCulture0403['KoreanCulture0403']=A_KoreanCulture0403['KoreanCulture0403'] * 60
        
        '''
        A_DancingCourse0406
        '''
        A_DancingCourse0406 = pd.read_excel('D:\Mass\Sports_data\Active\Courses\A_DancingCourse0406.xlsx')
        A_DancingCourse0406 =A_DancingCourse0406.drop_duplicates(subset = None, keep = 'last')
        A_DancingCourse0406.rename(columns={A_DancingCourse0406.columns[0]: "Student_ID" }, inplace=True)
        A_DancingCourse0406.rename(columns={A_DancingCourse0406.columns[1]:"DancingCourse0406_In"}, inplace =True)
        A_DancingCourse0406.rename(columns={A_DancingCourse0406.columns[2]:"DancingCourse0406_Out"}, inplace =True)
        
        A_DancingCourse0406['DancingCourse0406']=A_DancingCourse0406['DancingCourse0406_Out']-A_DancingCourse0406['DancingCourse0406_In']
        A_DancingCourse0406['DancingCourse0406']=A_DancingCourse0406['DancingCourse0406'] / np.timedelta64(60,'s')
        
        '''
        A_DancingCourse0413
        '''
        A_DancingCourse0413 = pd.read_excel('D:\Mass\Sports_data\Active\Courses\A_DancingCourse0406.xlsx')
        A_DancingCourse0413 =A_DancingCourse0413.drop_duplicates(subset = None, keep = 'last')
        A_DancingCourse0413.rename(columns={A_DancingCourse0413.columns[0]: "Student_ID" }, inplace=True)
        A_DancingCourse0413.rename(columns={A_DancingCourse0413.columns[1]:"DancingCourse0413_In"}, inplace =True)
        A_DancingCourse0413.rename(columns={A_DancingCourse0413.columns[2]:"DancingCourse0413_Out"}, inplace =True)
        
        A_DancingCourse0413['DancingCourse0413']=A_DancingCourse0413['DancingCourse0413_Out']-A_DancingCourse0413['DancingCourse0413_In']
        A_DancingCourse0413['DancingCourse0413']=A_DancingCourse0413['DancingCourse0413'] / np.timedelta64(60,'s')
        
        '''
        A_DancingCourse0427
        '''
        A_DancingCourse0427 = pd.read_excel('D:\Mass\Sports_data\Active\Courses\A_DancingCourse0406.xlsx')
        A_DancingCourse0427 =A_DancingCourse0427.drop_duplicates(subset = None, keep = 'last')
        A_DancingCourse0427.rename(columns={A_DancingCourse0427.columns[0]: "Student_ID" }, inplace=True)
        A_DancingCourse0427.rename(columns={A_DancingCourse0427.columns[1]:"DancingCourse0427_In"}, inplace =True)
        A_DancingCourse0427.rename(columns={A_DancingCourse0427.columns[2]:"DancingCourse0427_Out"}, inplace =True)
        
        A_DancingCourse0427['DancingCourse0427']=A_DancingCourse0427['DancingCourse0427_Out']-A_DancingCourse0427['DancingCourse0427_In']
        A_DancingCourse0427['DancingCourse0427']=A_DancingCourse0427['DancingCourse0427'] / np.timedelta64(60,'s')
        
        '''
        A_SwimTime
        '''
        A_SwimTime = pd.read_excel('D:\Mass\Sports_data\Active\Courses\A_SwimTime.xlsx')
        A_SwimTime =A_SwimTime.drop_duplicates(subset = None, keep = 'last')
        A_SwimTime.rename(columns={A_SwimTime.columns[0]: "Student_ID" }, inplace=True)
        A_SwimTime.rename(columns={A_SwimTime.columns[1]:"SwimTime"}, inplace =True)
        A_SwimTime['SwimTime']=A_SwimTime['SwimTime'] * 60
        
        '''
        SwimBeginner
        '''
        SwimBeginner = pd.read_excel('D:\Mass\Sports_data\Active\Courses\A_IC_01SwimBeginner.xlsx')
        SwimBeginner =SwimBeginner.drop_duplicates(subset = None, keep = 'last')
        SwimBeginner.rename(columns={SwimBeginner.columns[0]: "Student_ID" }, inplace=True)
        SwimBeginner.rename(columns={SwimBeginner.columns[1]:"SwimBeginner"}, inplace =True)
        SwimBeginner['SwimBeginner']=SwimBeginner['SwimBeginner'] * 60
        
        '''
        WingChun
        '''
        WingChun = pd.read_excel('D:\Mass\Sports_data\Active\Courses\A_IC_02WingChun.xlsx')
        WingChun =WingChun.drop_duplicates(subset = None, keep = 'last')
        WingChun.rename(columns={WingChun.columns[0]: "Student_ID" }, inplace=True)
        WingChun.rename(columns={WingChun.columns[1]:"WingChun"}, inplace =True)
        WingChun['WingChun']=WingChun['WingChun'] * 60
        
        '''
        A_IC_03Yoga
        '''
        Yoga = pd.read_excel('D:\Mass\Sports_data\Active\Courses\A_IC_03Yoga.xlsx')
        Yoga =Yoga.drop_duplicates(subset = None, keep = 'last')
        Yoga.rename(columns={Yoga.columns[0]: "Student_ID" }, inplace=True)
        Yoga.rename(columns={Yoga.columns[1]:"Yoga"}, inplace =True)
        Yoga['Yoga']=Yoga['Yoga'] * 60
        
        '''
        AerialYoga
        '''
        AerialYoga = pd.read_excel('D:\Mass\Sports_data\Active\Courses\A_IC_04AerialYoga.xlsx')
        AerialYoga =AerialYoga.drop_duplicates(subset = None, keep = 'last')
        AerialYoga.rename(columns={AerialYoga.columns[0]: "Student_ID" }, inplace=True)
        AerialYoga.rename(columns={AerialYoga.columns[1]:"AerialYoga"}, inplace =True)
        AerialYoga['AerialYoga']=AerialYoga['AerialYoga'] * 60
        
        
        '''
        Kickboxing
        '''
        Kickboxing = pd.read_excel('D:\Mass\Sports_data\Active\Courses\A_IC_06Kickboxing.xlsx')
        Kickboxing =Kickboxing.drop_duplicates(subset = None, keep = 'last')
        Kickboxing.rename(columns={Kickboxing.columns[0]: "Student_ID" }, inplace=True)
        Kickboxing.rename(columns={Kickboxing.columns[1]:"Kickboxing"}, inplace =True)
        Kickboxing['Kickboxing']=Kickboxing['Kickboxing'] * 60
        
        '''
        Squash
        '''
        Squash = pd.read_excel('D:\Mass\Sports_data\Active\Courses\A_IC_07Squash.xlsx')
        Squash =Squash.drop_duplicates(subset = None, keep = 'last')
        Squash.rename(columns={Squash.columns[0]: "Student_ID" }, inplace=True)
        Squash.rename(columns={Squash.columns[1]:"Squash"}, inplace =True)
        Squash['Squash']=Squash['Squash'] * 60
        
        '''
        RopeSkipping
        '''
        RopeSkipping = pd.read_excel('D:\Mass\Sports_data\Active\Courses\A_IC_08RopeSkipping.xlsx')
        RopeSkipping =RopeSkipping.drop_duplicates(subset = None, keep = 'last')
        RopeSkipping.rename(columns={RopeSkipping.columns[0]: "Student_ID" }, inplace=True)
        RopeSkipping.rename(columns={RopeSkipping.columns[1]:"RopeSkipping"}, inplace =True)
        RopeSkipping['RopeSkipping']=RopeSkipping['RopeSkipping'] * 60
        
        '''
        Taekwonodo
        '''
        Taekwonodo = pd.read_excel('D:\Mass\Sports_data\Active\Courses\A_IC_09Taekwonodo.xlsx')
        Taekwonodo =Taekwonodo.drop_duplicates(subset = None, keep = 'last')
        Taekwonodo.rename(columns={Taekwonodo.columns[0]: "Student_ID" }, inplace=True)
        Taekwonodo.rename(columns={Taekwonodo.columns[1]:"Taekwonodo"}, inplace =True)
        Taekwonodo['Taekwonodo']=Taekwonodo['Taekwonodo'] * 60
        
        
        
        '''
        数据整合
        '''
        #表的关联
        ##比赛
        Reletive = Student_info.join(A_BasketGame0309.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_BasketGame0310.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_BasketGame0316.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_BasketGame0317.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_Badminton.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_Swim0219.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_Swim0226.set_index('Student_ID'),on='Student_ID',how='left')
        
        Reletive = Reletive.join(A_Soccer0330.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_Soccer0331.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_Squash0330.set_index('Student_ID'),on='Student_ID',how='left')
        
        
        Reletive = Reletive.join(A_TableTennis.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(TableTennisClub.set_index('Student_ID'),on='Student_ID',how='left')
        
        
        Reletive = Reletive.join(A_TrackField.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_Volleyball.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_Volleyball0323.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_Volleyball0323PM.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_Volleyball0324.set_index('Student_ID'),on='Student_ID',how='left')
       
          
        Reletive = Reletive.join(Badminton_Un.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(DragonDance.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(BasketballMen.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(BasketballWomenCD.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(BasketballMenCD.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(BasketballWomen.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(Fencing.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(Opening0427.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(Competi0427.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(Dancing0413.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(Dancing0424.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(FlagRaising.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(Flag0430.set_index('Student_ID'),on='Student_ID',how='left')

        Reletive = Reletive.join(ClimbingSession1.set_index('Student_ID'),on='Student_ID',how='left')        
        Reletive = Reletive.join(ClimbingSession2.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(HKUST_UM_Cup.set_index('Student_ID'),on='Student_ID',how='left')        
        Reletive = Reletive.join(FSTBball.set_index('Student_ID'),on='Student_ID',how='left')        
        Reletive = Reletive.join(FSTFutsal.set_index('Student_ID'),on='Student_ID',how='left')        
        Reletive = Reletive.join(A_UMOpening.set_index('Student_ID'),on='Student_ID',how='left')        
        
        #coursess
        Reletive = Reletive.join(A_DiaboloCourse0228.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_DiaboloCourse0307.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_DiaboloCourse0314.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_DiaboloCourse0321.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_DiaboloCourse0328.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_WomenSoccor.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_Tennis0320.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_Tennis0327.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_TableTennis04.set_index('Student_ID'),on='Student_ID',how='left')
        
        Reletive = Reletive.join(A_UMSUTrack_Field.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_SwimTime.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_SwimTime04.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_TennisInterestCourse.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_DancingCourse0406.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_DancingCourse0413.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_DancingCourse0427.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_KoreanCulture0309.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_KoreanCulture0316.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_KoreanCulture0323.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_KoreanCulture0330.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_KoreanCulture0403.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_Korfball0413.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_Korfball0420.set_index('Student_ID'),on='Student_ID',how='left')
        
        Reletive = Reletive.join(SwimBeginner.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(WingChun.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(Yoga.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(AerialYoga.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(Kickboxing.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(Squash.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(RopeSkipping.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(Taekwonodo.set_index('Student_ID'),on='Student_ID',how='left')
        
        
        
        #Runing
        Reletive = Reletive.join(A_JunKR0108.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_JunKR0110.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_JunKR0117.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_JunKR0122.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_JunKR0124.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_JunKR0129.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_NightRun0117.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_NightRun0228.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_NightRun0404.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_NightRun0416.set_index('Student_ID'),on='Student_ID',how='left')     

        Reletive = Reletive.join(A_NightWalk.set_index('Student_ID'),on='Student_ID',how='left')
        
        Reletive = Reletive.join(A_KeepRun0305.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_KeepRun0312AM.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_KeepRun0312PM.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_KeepRun0313.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_KeepRun0314.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_KeepRun0319AM.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_KeepRun0319PM.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_KeepRun0320.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_KeepRun0321.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_KeepRun0326AM.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_KeepRun0326PM.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_KeepRun0327.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_KeepRun0328.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_Letexercise_session1.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_Letexercise_session2.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_Letexercise_session3.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_Letexercise_session4.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_Letexercise_session5.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_Letexercise_session6.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_Letexercise_session7.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_Letexercise_session8.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_Letexercise_session9.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_Letexercise_session10.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_Letexercise_session11.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_Letexercise_session12.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_Letexercise_session13.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_Letexercise_session14.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_Letexercise_session15.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_Letexercise_session16.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_Letexercise_session17.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_Letexercise_session18.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_Letexercise_session19.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_Letexercise_session20.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_Letexercise_session21.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_Letexercise_session22.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_Letexercise_session23.set_index('Student_ID'),on='Student_ID',how='left')
        Reletive = Reletive.join(A_Letexercise_session24.set_index('Student_ID'),on='Student_ID',how='left')
        
        Reletive= Reletive.fillna(0)
        
        Reletive['SwimCompetition']=Reletive[['Swim0126','Swim0119']].sum(axis=1)
        Reletive['BasketCompetition']=Reletive['BasketGame0310']+Reletive['BasketGame0309']
        Reletive['BasketCompetition2']=Reletive['BasketGame0316']+Reletive['BasketGame0317']
        Reletive['BasketTeamCompetition']=Reletive[['BasketballMen','BasketballMenCD','BasketballWomen','BasketballWomenCD']].sum(axis=1)
        #Reletive['Dancing']=Reletive[['DancingTongXin','DancingNianLun']].sum(axis=1)
        Reletive['ClimbingCompetition']=Reletive[['ClimbingSession1','ClimbingSession2']].sum(axis=1)
        #Reletive['FSTCup']=Reletive[['FSTBball','FSTFutsal']].sum(axis=1)
        Reletive['VolleyballCompetition']=Reletive[['A_Volleyball','A_Volleyball0323','A_Volleyball0323PM','A_Volleyball0324']].sum(axis=1)
        Reletive['Soccor_Competition']=Reletive[['Soccer0330','Soccer0331']].sum(axis=1)
        Reletive['Letsplay_TableTennis']=Reletive[['Tennis0327','Tennis0320','TableTennis04']].sum(axis=1)
        Reletive['DiaboloCourse']=Reletive[[ 'DiaboloCourse0228','DiaboloCourse0307','DiaboloCourse0314',
                                            'DiaboloCourse0321','DiaboloCourse0328']].sum(axis=1)
        Reletive['DanceCourse']=Reletive[[ 'DancingCourse0406','DancingCourse0413','DancingCourse0427']].sum(axis=1)
        Reletive['KoreanCulture']=Reletive[[ 'KoreanCulture0309','KoreanCulture0316','KoreanCulture0323',
                                            'KoreanCulture0330','KoreanCulture0403']].sum(axis=1)
        Reletive['KorfballCourse']=Reletive[[ 'Korfball0413','Korfball0420']].sum(axis=1)
       
        Reletive['Letexercise']=Reletive[['Letexercise_session1','Letexercise_session2',
                                       'Letexercise_session3','Letexercise_session4','Letexercise_session5',
                                       'Letexercise_session6','Letexercise_session7','Letexercise_session8',
                                       'Letexercise_session9','Letexercise_session10','Letexercise_session11',
                                       'Letexercise_session12','Letexercise_session13','Letexercise_session14',
                                       'Letexercise_session15','Letexercise_session16','Letexercise_session17',
                                       'Letexercise_session18','Letexercise_session19','Letexercise_session20',
                                       'Letexercise_session21','Letexercise_session22','Letexercise_session23',
                                       'Letexercise_session24']].sum(axis=1)
        Reletive['Run']=Reletive[['KeepRun0108','KeepRun0110','KeepRun0117','KeepRun0122',
                                        'KeepRun0124','KeepRun0129','KeepRun0305','KeepRun0312AM',
                                        'KeepRun0312PM','KeepRun0313','KeepRun0314','KeepRun0319AM',
                                        'KeepRun0319PM','KeepRun0326AM','KeepRun0326PM','KeepRun0320',
                                        'KeepRun0321','KeepRun0327','KeepRun0328','NightWalk',                            
                                        'NightRun0117','NightRun0228','NightRun0404','NightRun0416']].sum(axis=1)
        Reletive['SwimTime']=Reletive[['SwimTime','SwimTime04']].sum(axis=1)  
        Reletive['Total']=Reletive[[
             # B
             'BasketCompetition', 'BasketCompetition2', 'BasketballMen',
             'BasketballMenCD', 'BasketballWomen', 'BasketballWomenCD',
             'BadmintonCompetition', 'Badminton_Un_Competition',
             # C
             'ClimbingCompetition', 'Competi_AoKeBei0427',
             # D
             'DancingTongXin', 'DancingNianLun', 'DragonDance',
             # H
             'HKUST_UM_Cup',
             # F
             'Flag54Ceremony', 'FencingCompetition', 'FlagRaising', 'FSTBball', 'FSTFutsal',
             # O
             'Opening_UM', 'Opening_AoKeBei0427',
             # S
             'SwimCompetition', 'Soccor_Competition', 'Squash_Competition',
             # T
             'TrackFieldCompetition', 'TableTennisCompetition', 'TableTennisClubCompetition',
             # V
             'VolleyballCompetition',

             'AerialYoga',
             'DiaboloCourse', 'DanceCourse',
             'Kickboxing', 'KoreanCulture', 'KorfballCourse',
             'Letsplay_TableTennis',
             'RopeSkipping',
             'SwimBeginner', 'SwimTime', 'Squash',
             'Taekwonodo', 'TennisInterestCourse',
             'UMSUTrack_Field',
             'WomenSoccor', 'WingChun',
             'Yoga',

             'Letexercise', 'Run'
        ]].sum(axis=1)
        Reletive_stu = Reletive.loc[:,["Student_ID","Eng_Name","Ch_Name",
                                       # B
                                       "BasketCompetition", "BasketCompetition2", "BasketballMen",
                                       "BasketballMenCD", "BasketballWomen", "BasketballWomenCD",
                                       "BadmintonCompetition", "Badminton_Un_Competition",
                                       # C
                                       "ClimbingCompetition", "Competi_AoKeBei0427",
                                       # D
                                       "DancingTongXin", "DancingNianLun", "DragonDance",
                                       # H
                                       "HKUST_UM_Cup",
                                       # F
                                       "Flag54Ceremony", "FencingCompetition", "FlagRaising", "FSTBball", "FSTFutsal",
                                       # O
                                       "Opening_UM", "Opening_AoKeBei0427",
                                       # S
                                       "SwimCompetition", "Soccor_Competition", "Squash_Competition",
                                       # T
                                       "TrackFieldCompetition", "TableTennisCompetition", "TableTennisClubCompetition",
                                       # V
                                       "VolleyballCompetition",

                                       "AerialYoga",
                                       "DiaboloCourse", "DanceCourse",
                                       "Kickboxing", "KoreanCulture", "KorfballCourse",
                                       "Letsplay_TableTennis",
                                       "RopeSkipping",
                                       "SwimBeginner", "SwimTime", "Squash",
                                       "Taekwonodo", "TennisInterestCourse",
                                       "UMSUTrack_Field",
                                       "WomenSoccor", "WingChun",
                                       "Yoga",

                                       "KeepRun0108","KeepRun0110","KeepRun0117","KeepRun0122","KeepRun0124","KeepRun0129",
                                       "KeepRun0305","KeepRun0312AM","KeepRun0312PM","KeepRun0313","KeepRun0314","KeepRun0319AM","KeepRun0319PM","KeepRun0326AM",
                                       "KeepRun0326PM","KeepRun0320","KeepRun0321","KeepRun0327","KeepRun0328",

                                       "Letexercise_session1","Letexercise_session2",
                                       "Letexercise_session3","Letexercise_session4","Letexercise_session5",
                                       "Letexercise_session6","Letexercise_session7","Letexercise_session8",
                                       "Letexercise_session9","Letexercise_session10","Letexercise_session11",
                                       "Letexercise_session12","Letexercise_session13","Letexercise_session14",
                                       "Letexercise_session15","Letexercise_session16","Letexercise_session17",
                                       "Letexercise_session18","Letexercise_session19","Letexercise_session20",
                                       "Letexercise_session21","Letexercise_session22","Letexercise_session23",
                                       "Letexercise_session24",

                                       "NightRun0117", "NightRun0228", "NightRun0404", "NightRun0416",
                                       "NightWalk",
                                       
                                       'Total'
                                       ]]
        
       
        
        #格式调整
        Reletive_stu = Reletive_stu.round(1)
        
        '''
        暂储数据
        '''
        #Reletive.to_excel("D:\Mass\Sports_data\Combine2019S1\CPED011.xlsx")
        Combine_file_name=Combine_dir +'/' + file.split('.')[0]+'.xlsx'
        Reletive_stu.to_excel(Combine_file_name,index=False)
        #Reletive_stu.to_excel("D:\Mass\Sports_data\Combine2019S1\Stu\CPED022.xlsx",index=False)
