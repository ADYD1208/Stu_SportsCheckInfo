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
```python
  '''
  学生数据导入
  '''
  Student_info = pd.read_excel('D:\Mass\Sports_data\Water_class\section019.xlsx')
```
### 添加活动
---
#### 1 添加表格
---
在**活动数据**模块中，添加如下代码行
```python 
'''
A_NightRun20190117
'''
      
A_NightRun =pd.read_excel('D:\Mass\Sports_data\Active\A_NightRun20190117.xlsx')         #导入表格，更改文件名，和读入表格地址
A_NightRun.rename(columns={A_NightRun.columns[0]:"Student_ID_Number"}, inplace=True)    #更改文件名*2
A_NightRun.rename(columns={A_NightRun.columns[1]:"NightRun_In"}, inplace=True )         #更改文件名*2，更改Check_in列名
A_NightRun.rename(columns={A_NightRun.columns[2]:"NightRun_Out"},inplace=True)          #更改文件名*2，更改Check_out列名

A_NightRun['NightRun']=A_NightRun['NightRun_Out']-A_NightRun['NightRun_In']             #更改文件名*2，活动列名更改，check in/out更改列名，注意 顺序别写反了
A_NightRun['NightRun']=A_NightRun['NightRun'] / np.timedelta64(60,'s')                  #更改文件名*2，活动列名更改
```
---
以添加活动为UM Opening为例，其中需要更改的变量名有：
* 将A_NightRun改为相对应活动名称。例如，添加UMOpening为活动，则将所有变量A_NightRun改为A_UMOpening
* 同时修改需要导入的表格以及表格地址
* 更改签到表列名，Check in & Check out列名改为相对应活动列名。例如，将夜跑活动的NightRun，改为UMOpening;将夜跑活动的NightRun_In，改为Opening_In; 将将夜跑活动的NightRun_Out，改为Opening_Out（注意要改两次）
---
更改后的代码为
```python
'''
A_UMOpening20181024
'''
A_UMOpening = pd.read_excel('D:\Mass\Sports_data\Active\A_UM Sports Day Opening.xlsx')
A_UMOpening.rename(columns={A_UMOpening.columns[0]: "Student_ID_Number" }, inplace=True)
A_UMOpening.rename(columns={A_UMOpening.columns[1]:"UMOpening_In"}, inplace =True)
A_UMOpening.rename(columns={A_UMOpening.columns[2]:"UMOpening_Out"}, inplace =True)

A_UMOpening['UMOpening']=A_UMOpening['UMOpening_Out']-A_UMOpening['UMOpening_In']
A_UMOpening['UMOpening']=A_UMOpening['UMOpening'] / np.timedelta64(60,'s')

```
#### 2 添加表格关联
---
在表的关联部分添加一行代码
```python
Reletive = Reletive.join(A_NightRun.set_index('Student_ID_Number'),on='Student_ID_Number',how='left')
```
还以添加活动UMOpening为例更改文件名与刚才所更改的活动文件名一致，如A_NightRun改为A_UMOpening
```python
Reletive = Reletive.join(A_UMOpening.set_index('Student_ID_Number'),on='Student_ID_Number',how='left')
```
#### 3 修改某班级对应n个活动的表格存储信息
---
更改所需存储信息的位置和表格名称，注意区分班级
```
Reletive.to_excel("D:\Mass\Sports_data\Combine\CPED1002_019.xlsx")
```
## 常见错误
当列名，例如上述代码中的```NightRun```，```NightRun_In```和```NightRun_Out```没有完全修改成功时，会出现列名重复的报错。
```
ValueError: columns overlap but no suffix specified: Index(['Diabolo1024'], dtype='object')
```
