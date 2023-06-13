import pandas as pd
import datetime
import time
import sys
import openpyxl

print("reading file....")
#读取需要操作的数据进来
updateData= pd.read_excel("update_info.xlsx")
origin_Data = pd.read_excel("DAR 0605.xlsx")
f=open("log.txt","+a",encoding="UTF-8") 

print("Farmat adjusting...")
#格式调整
original_title = origin_Data.columns
new_title = list(origin_Data.iloc[0])
origin_Data.columns = new_title
DF = origin_Data.drop(labels=0,axis=0)
DF = DF.sort_values(by=["ID#"])
#数据格式调整
for i in updateData.columns:
    if updateData[i].dtype == "object":
        pass
    else:
        updateData[i] = updateData[i].astype("object")

#找到学校信息更新表是哪个学校的！
school_list = ['DCB','DCSG','DCSL','DCSPD','DCSPX','DCSZ','DEGT','DEMH','DEXA','DHSZ','DHZH','DUSZ','HQ']
school_name =""
for i in school_list:
    if updateData[i][1] == "Y":
        school_name = i

#定义错误类
class Error(Exception):
    def __init__(self, message):
        self.message = message

    def __str__(self):
    	# 异常抛出后的日志信息，对应raise HaveValueError中输出
        return self.message

print("Matching each columns...")
#DTA in place 拆分找到对应列
DF_col = list(DF.columns)
updateData_col = list(updateData.columns)
updateData_col_new = []
for i in updateData_col:
    try :
        DF_col.index(i)
    except ValueError as v:
        if str(v) == "\'DTA in place\' is not in list":
            DTA = "DTA "+school_name
            updateData_col_new.append(DTA) 
        else:
            print("can't match column “"+ i+ "” with the table,Please make sure the Columns was correspond to each other")
            time.sleep(5)
            sys.exit()
            
    else:
        updateData_col_new.append(i)
updateData.columns = updateData_col_new


#为需要新创建的产品ID值序号。找到最后一个产品的ID值
a = list(DF["ID#"])
index = []
for k in a:
    index.append(int(k[1:]))
index.sort()
index_new = index[-3]

now_time = datetime.datetime.now()
Log_Time = "Time: "+str(now_time) +"\n"
f.write("---------------------------------------------------------------------------------------------------------------\n")
f.write(Log_Time)
Log_title = "School/System: "+school_name +"\n"
f.write(Log_title)
f.write("更新内容：\n")


print("Updating....")
#遍历反馈与信息表：
for i in updateData.iterrows():
    #找到产品名称
    name = i[1]["Product / Service Name"]
    
    #构建新产品的数据字典（原则是只添不删）
    data_dict = dict(i[1])
    new_dataDict = {}
    
    for k,v in data_dict.items():
        if str(v) != "nan":
            new_dataDict[k] = v
        if k == "Data location":
            if str(v) == "nan" or str(v).upper() == "OUTSIDE FIREWALL":
                new_dataDict["Data location"] = "Outside Firewall"
            else:
                new_dataDict["Data location"] = v
            
     #判断更新数据中的datalocation字段是否有值       
    if new_dataDict["Data location"] == "Outside Firewall":
        #判断该产品是否存在
        #如果不存在，则创建新的一条的产品信息
        if len(DF[DF["Product / Service Name"]==name]) == 0:
            #准备ID值
            index_new+=1
            count0 = 6-len(str(index_new + 1))
            ID = "#"+count0*"0"+str(index_new)
            new_dataDict["ID#"] = ID
            #把新行添加到大表中
            DF_ = pd.DataFrame([new_dataDict])
            DF = pd.merge(DF,DF_,how="outer")
            Log = name+" 是新产品，添加到大表中，ID为："+ID+"，Data Location为:"+new_dataDict["Data location"]+"\n"
            f.write(Log)
         #如果存在，则更新大表中除了datalocation以为的该行的信息
        else:
            
            DF_ = DF.loc[(DF["Product / Service Name"] == name) & (DF[school_name] == 'Y')]
            idx = DF_.index
            DF = DF.drop(idx)                
            for k,v in new_dataDict.items():
                if k == "Data location":
                    pass
                else:
                    DF_[k] = v
            DF = pd.concat([DF,DF_])
            Log = name+" 数据已经存在，且Data Location没有发生变化，只需要更新原表其他内容"+"\n"
            f.write(Log)
    else:
        #判断该产品是否存在
        #如果不存在，则创建新的一条的产品信息
        if len(DF[DF["Product / Service Name"]==name]) == 0:
            #准备ID值
            index_new+=1
            count0 = 6 - len(str(index_new + 1))
            ID = "#"+count0*"0"+str(index_new)
            new_dataDict["ID#"] = ID
            #把新行添加到大表中
            DF_ = pd.DataFrame([new_dataDict])
            DF = pd.merge(DF,DF_,how="outer")
            Log = name+" 是新产品，添加到大表中，ID为："+ ID + "，Data Location为:"+new_dataDict["Data location"]+"\n"
            f.write(Log) 
        #如果存在
        else:
            #判断大表中的datalocation和更新的datalocation是否一致：
            DF_ = DF.loc[(DF["Product / Service Name"] == name) & (DF[school_name] == 'Y')]
            idx = DF_.index
            DLocation = DF_["Data location"].values[0]
            if DLocation == new_dataDict["Data location"]:
                #如果一致,则更新其中的信息。
                DF = DF.drop(idx)
                for k,v in new_dataDict.items():
                    if k == "Data location":
                        pass
                    else:
                        DF_[k] = v
                DF = pd.concat([DF,DF_])
                Log = name+" 数据已经存在，且Data Location没有发生变化，只需要更新原表其他内容"+"\n"
                f.write(Log)
                
            else:
                #如果不一致，则将该行中的学校选项去掉，并新增一行，且生成新的一行。
                #对原始行进行操作,去掉该学校勾选。

                #判断datalocation是否是从0到1
                #如果是从没有（NAN/Outside Firewall）到具体的某个地方，无需重新起一行
                if DLocation =="nan" or str(DLocation).upper() =="OUTSIDE FIREWALL":
                    DF =DF.drop(idx)
                    for k,v in new_dataDict.items():
                            DF_[k] = v
                    DF = pd.concat([DF,DF_])
                    Log = name+" 数据已存在！Data Location由“Outside Firewall”变更为了:"+new_dataDict["Data location"]+"\n"
                    f.write(Log)
                else:
                    #如果是从一个地方到另一个地方则需要重起一行


                    # 判断是否多个学校共有一条数据
                    count = 0
                    for i in school_list:
                        if DF_[i].values[0] =="Y":
                            count +=1
                    
                    if count <=1:
                        #只有一个学校使用该条数据，无需重新起一行，只需要修改datalocation
                        DF = DF.drop(idx)
                        for k,v in new_dataDict.items():
                            DF_[k] = v
                        DF = pd.concat([DF,DF_])
                        Log = name +" 数据已存在！只有一个学校/系统在使用该条数据，Data Location由“"+DLocation+"”变更为了:"+new_dataDict["Data location"]+"，无需重新再起一行\n"
                        f.write(Log)
                    else:
                        #多个学校共用一条数据
                        DF_1 = DF_
                        DF_2 = DF_
                        DF_1[school_name] = "N"
                        DF = DF.drop(idx)
                        DF = pd.concat([DF,DF_1])
                        DF = DF.sort_values(by=["ID#"])
                        #添加新的行
                        index_new+=1
                        count0 = 6 - len(str(index_new + 1))
                        ID = "#"+count0*"0"+str(index_new)
                        new_dataDict["ID#"] = ID
                        #把新行添加到大表中，更改datalocation，并去除其他学校的勾选行
                        for i in school_list:
                            if i == school_name:
                                DF_2[i] = "Y"
                            else:
                                DF_2[i] = "N"
                        for k ,v in new_dataDict.items():
                            DF_2[k] = v 
                        DF = pd.concat([DF,DF_2])
                        DF = DF.sort_values(by=["ID#"])
                        Log = name+" 与原有的Data Location不一致,且多个系统/学校公用该条数据,由“"+DLocation+"”变更为了:"+new_dataDict["Data location"]+",另起一行，ID是："+ID+"\n"
                        f.write(Log)
f.write("---------------------------------------------------------------------------------------------------------------\n")
f.write("\n")                
f.close()
DF = DF.sort_values(by=["ID#"])
DF.to_excel("test.xlsx",index=False)

print("Successfully!")
time.sleep(2)

#test\