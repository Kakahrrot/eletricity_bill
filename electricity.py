import pandas
import re

name = pandas.read_excel("別名對照.xlsx", na_filter=False)
path = None
yearNum = None
table = None

def init_elec(Path, YearNum):
    global path, yearNum, table
    path = Path
    yearNum = YearNum
    df = pandas.read_excel(path+"110年度系所電費統計表.xls", sheet_name = "電費統計表",na_filter=False)
    df.columns = df.columns.str.replace(" ", "")

    table = pandas.DataFrame()
    table["使用單位"] = df["使用單位"]

    for i in range(0, 4):
        table[str(yearNum - i)+"實際用電度數"] = df[str(yearNum - i)+"實際用電度數"]
    table["近三年平均使用度數\n(b)"] = df["近三年平均使用度數\n(b)"]

def search(df, col, item):
    found = False
    res = df.loc[df["單位"] == item][col]
    if len(res)!= 0:
        found = True
    else: # serach for alias
        alias = name.loc[name["原始名稱"] == item]["別名"]
        if len(alias) != 0:
            alias = alias.values[0].split()
            for k in alias:
                res = df.loc[df["單位"] == k][col]
                if len(res)!= 0:
                    found = True
                    break
    return found, res

def extract(file, col, col_name, sheet = 0, CENTER = False):
    print("\n\n", file, col_name)
#     df = pandas.read_excel(path+file, na_filter=False)
    df = pandas.read_excel(path+file, sheet_name=sheet, na_filter=False)

    df.columns = df.columns.str.replace(" ", "")
    df.columns = df.columns.str.replace("\n", "")
    df.columns = df.columns.str.replace("\t", "")

#     df["單位"] = df["單位"].str.replace("(", "", regex=False)
#     df["單位"] = df["單位"].str.replace(")", "", regex=False)
    df["單位"] = df["單位"].str.replace(" ", "", regex=False)
    df["單位"] =  df['單位'].str.replace(r"\(.*\)","", regex=True)
    df["單位"] =  df['單位'].str.replace(r"（.*）","", regex=True)

    # loop through table and extract data from df
    result = []

    for i in table["使用單位"]:
        if "※計算參數" in i:
            result.append(0)
            continue
        searchList = i.replace("(", " ").replace(")", "").replace("、", " ").replace("含", " ").split()
    #     searchList.extend(name.loc[name["原始名稱"]== i]["別名"])
        searchList = [x for x in searchList if x != ""]
#         print(i, searchList)
        t = 0
        centerCounter = 0
        for j in searchList:
            found, res = search(df, col, j)
            if CENTER:
                if "中心" in j:
                    centerCounter +=1
                    if found:
#                         print(j, "found!!", res.values.sum(), col_name+str(centerCounter))
                        table.loc[table["使用單位"] == i, [col_name+str(centerCounter)]] = res.values.sum()
                    else:
                        print(j, "not found")
                        
            else:
                if "中心" in j:
                    t += 0
                else:
                    if found:
                        t += res.values.sum()
                    elif "總計" not in j and "中心" not in j and "學院" not in j and "開始計費" not in j:
                        print(j, "not found")
        result.append(t)
    if not CENTER:
        table[col_name] = result


def plan(file, col, col_name, sheet = None):
    print("\n\n", file, sheet, col_name)
    df = pandas.read_excel(path+file, sheet_name=sheet, na_filter=False)
    # skip the first line of the table and create new header

#     new_header = df.iloc[0] #grab the first row for the header
#     df = df[1:] #take the data less the header row
#     df.columns = new_header #set the header row as the df header
    df.columns = df.columns.str.replace(" ", "")
    df.columns = df.columns.str.replace("\n", "")
    df.columns = df.columns.str.replace("\t", "")
    df["單位"] = df["單位"].str.replace("[a-zA-Z0-9]+", "", regex=True)
    df["單位"] = df["單位"].str.replace("國立成功大學", "", regex=False)
    df["單位"] = df["單位"].str.replace("（所）", "", regex=False)

    # loop through table and extract data from df
    result0 = []
    result1 = []
#     print(df.loc[df["單位"] == "工業與資訊管理學系"][col])
#     return
    for i in table["使用單位"]:
        if "※計算參數" in i:
            result0.append(t0)
            result1.append(t1)
            continue
        searchList = i.replace("(", " ").replace(")", "").replace("、", " ").replace("含", " ").split()
    #     searchList.extend(name.loc[name["原始名稱"]== i]["別名"])
        searchList = [x for x in searchList if x != ""]
#         print(i, searchList)
        t0 = 0
        t1 = 0
        
        for j in searchList:
            found, res = search(df, col, j)
            if "中心" in j :
                t0 += 0
                t1 += 0
            else:
                if found:
                    t0 += res.loc[res["管理費"] != 0]["核定經費"].sum()
                    t1 += res.loc[res["管理費"] == 0]["核定經費"].sum()
                elif not found and "總計" not in j and "學院" not in j and "開始計費" not in j:
                    print(j, "not found")
        result0.append(t0)
        result1.append(t1)
    table[col_name[0]] = result0
    table[col_name[1]] = result1

def electricity(tt):
    # plan(file = "110年計畫（總）.xlsx", col = ["管理費","核定經費"],col_name = ["建教(有)", "建教(無)"],sheet = "建教")
    # extract(file = "110年度各級中心電費配額-3.xlsx", col = ["電費配額：「用電配額」乘「所有具有獨立電表中心之電費配額與用電配額比值」"],col_name = "中心", sheet = "（3）以產創中心方式-無獨立電表中心電費配額", CENTER = True)
    for data in tt:
        file = data[0]
        col = data[1].split()
        col_name = data[2].split()
        sheet = data[3]
        center = data[4]
        if center == "True":
            center = True
        if center == "False":
            center = False

        if len(col_name) > 1:
            plan(file = file, col = col, col_name = col_name, sheet = sheet if sheet != "" else 0)
        else:
            extract(file = file, col = col, col_name = col_name[0], sheet = sheet if sheet != "" else 0, CENTER = center)

    table.to_excel("output.xlsx") 