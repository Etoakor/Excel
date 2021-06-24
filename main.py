# -*- coding: utf-8 -*-


import pandas as pd

data = pd.read_excel("D:\excel\月报.xls")
rows = data.shape[0]  # 获取行数 shape[1]获取列数
department_list = []

for i in range(rows):
    temp = data["PROV_DESC"][i]
    if temp not in department_list:  # 防止重复
        department_list.append(temp)  # 将省分的分类存在一个列表中

n = len(department_list)  # 31省

df1 = pd.DataFrame()  # 用于各省的dataframe
df2 = pd.DataFrame()
df3 = pd.DataFrame()
df4 = pd.DataFrame()
df5 = pd.DataFrame()
df6 = pd.DataFrame()
df7 = pd.DataFrame()
df8 = pd.DataFrame()
df9 = pd.DataFrame()
df10 = pd.DataFrame()
df11 = pd.DataFrame()
df12 = pd.DataFrame()
df13 = pd.DataFrame()
df14 = pd.DataFrame()
df15 = pd.DataFrame()
df16 = pd.DataFrame()
df17 = pd.DataFrame()
df18 = pd.DataFrame()
df19 = pd.DataFrame()
df20 = pd.DataFrame()
df21 = pd.DataFrame()
df22 = pd.DataFrame()
df23 = pd.DataFrame()
df24 = pd.DataFrame()
df25 = pd.DataFrame()
df26 = pd.DataFrame()
df27 = pd.DataFrame()
df28 = pd.DataFrame()
df29 = pd.DataFrame()
df30 = pd.DataFrame()
df31 = pd.DataFrame()
df32 = pd.DataFrame()
df_list = [df1,df2,df3,df4,df5,df6,df7,df8,df9,df10,df11,df12,df13,df14,df15,df16,df17,df18,df19,df20,df21,df22,df23,df24,df25,df26,df27,df28,df29,df30,df31,df32]

for department in range(n):
    for i in range(0, rows):
        if data["PROV_DESC"][i] == department_list[department]:
            df_list[department] = pd.concat([df_list[department], data.iloc[[i], :]], axis=0, ignore_index=True)

writer = pd.ExcelWriter('D:/excel/5G用户规模月报.xls')  # 利用pd.ExcelWriter()存多张sheets

for i in range(n):
    df_list[i].to_excel(writer, sheet_name=str(department_list[i]), index=False)  # 加上index=FALSE 去掉index列

writer.save()