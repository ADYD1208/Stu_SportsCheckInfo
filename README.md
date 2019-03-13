# Stu_SportsCheckInfo
To collection the student information with their Check in &amp; Check out information of University.

## 环境
根据操作指南安装[Python 3.7.1](https://www.python.org/downloads/release/python-372/)和[Anaconda](https://medium.com/fishtung/python-anaconda-%E7%92%B0%E5%A2%83%E5%AE%89%E8%A3%9D%E6%95%99%E5%AD%B8-86bd13f8399d)

## 运行
打开Anaconda-Spyder
打开脚本File-open file 并运行Run
生成结果存于Variable Explorer中 Relative表中。

## 修改关联表
### 更换班级

将如下代码行替换成需要输入的学生班级表格
'''
学生数据导入
Student_info = pd.read_excel('D:\Mass\Sports_data\Water_class\section019.xlsx')
'''
### 添加活动
#### 1 添加表格
在-活动数据-模块中，添加如下代码行
'''
'''
A_NightRun20190117
'''
A_NightRun =pd.read_excel('D:\Mass\Sports_data\Active\A_NightRun20190117.xlsx')
A_NightRun.rename(columns={A_NightRun.columns[0]:"Student_ID_Number"}, inplace=True)
A_NightRun.rename(columns={A_NightRun.columns[1]:"NightRun_In"}, inplace=True )
A_NightRun.rename(columns={A_NightRun.columns[2]:"NightRun_Out"},inplace=True)

A_NightRun['NightRun']=A_NightRun['NightRun_Out']-A_NightRun['NightRun_In']
A_NightRun['NightRun']=A_NightRun['NightRun'] / np.timedelta64(60,'s')
'''
其中需要更改的变量名有：

#### 2 添加表格关联
