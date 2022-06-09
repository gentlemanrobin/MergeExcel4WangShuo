import pandas as pd
import os


# 打包方法：
# 以管理员身份运行CMD
# pyinstaller -F C:\Users\RobinGao\Desktop\Code\python\MergeExcel4WangShuo\mergeExcel.py --distpath=C:\Users\RobinGao\Desktop\Code\python\MergeExcel4WangShuo\dist
def get_sheet_previous_name(sheetName):
    """
    拼接sheet_previous_name的名字
    e.g. RE0505#02

    :return: sheet_previous_name
    """
    sheet_previous_name = sheetName.replace(".xlsx", "")
    print("准备处理：" + sheet_previous_name)
    # 下面对文件名做初步处理，用find查找是否包含#
    # 如果包含就往前取6个字符（批号）往后取2个字符（片号）,如果不包含会返回-1
    targetCharIndex = sheet_previous_name.find('#')
    if targetCharIndex != -1 and len(sheet_previous_name) >= 9:
        name1 = sheet_previous_name[targetCharIndex - 6:targetCharIndex]
        name2 = sheet_previous_name[targetCharIndex + 1:targetCharIndex + 3]
        sheet_previous_name = name1 + "#" + name2
    else:
        print("文件名不符合规则")

    print("处理完毕名字前缀=" + sheet_previous_name)
    return sheet_previous_name


def get_sheet_name(sheetName):
    """
    这个方法是通过for循环文件名来拼接sheet1和2的名字
     e.g. RE0505#02_CPRT RE0505#02_CPHT

    :return: excelSheetName
    """
    if 'CPH' in d2Name:
        sheetName = get_sheet_previous_name(sheetName) + "_CPHT"
    elif 'CPR' in d2Name:
        sheetName = get_sheet_previous_name(sheetName) + "_CPRT"
    print("sheetName==" + sheetName)
    return sheetName


if __name__ == '__main__':
    desktopPath = os.path.join(os.path.expanduser('~'), "Desktop")
    inputFilePath = desktopPath + r'\MERGE'
    outputFilePath = desktopPath + r'\AE7252#05-COMP1.xlsx'
    if os.path.exists(r'%s' % outputFilePath):
        os.remove(r'%s' % outputFilePath)
    # 创建一个目标文件,如果直接用文件路径，那么每次都会覆盖写入，ExcelWriter可以看作一个容器，一次性提交所有to_excel语句后再保存，
    # 从而避免覆盖写入。其实这一句代码只是创建一个新的XLSX文件。
    result = pd.ExcelWriter(r'%s' % outputFilePath)
    origin_file_list = os.listdir(r'%s' % inputFilePath)
    # 定义4个临时存储文件内容变量,如果后期文件多，要考虑用冒泡排序，而不是手动to_excel
    contentCPRT = ""
    contentCPHT = ""
    excelSheetNameCPRT = ""
    excelSheetNameCPHT = ""

    # 循环初始文件，写入目标xlsx文件
    for i in origin_file_list:
        file_path = r'%s' % inputFilePath + "\\" + i
        content = pd.read_excel(file_path, header=None)
        # 取单元格的值,d2的意思是表格的第二行第四列
        d2Name = content.loc[1][3]
        print("单元格名字:" + d2Name)
        # 这个地方为了方便强制写死了，如果文件数量多要考虑冒泡排序算法。
        if 'CPH' in d2Name:
            excelSheetNameCPHT = get_sheet_name(i)
            contentCPHT = content
        elif 'CPR' in d2Name:
            excelSheetNameCPRT = get_sheet_name(i)
            contentCPRT = content

    # index=False代表不添加索引
    contentCPRT.to_excel(result, excelSheetNameCPRT, index=False, header=None)
    contentCPHT.to_excel(result, excelSheetNameCPHT, index=False, header=None)
    df = pd.DataFrame()
    df.to_excel(result, get_sheet_previous_name(origin_file_list[0]) + "_COMP")
    df.to_excel(result, "RT_fail")
    df.to_excel(result, "HT_fail")
    df.to_excel(result, "良品数量")
    # 保存上面的操作，因为下面要读取最新的表格
    result.save()
    # 重新读取，当前的ExcelWriter支持覆盖写入
    result = pd.ExcelWriter(r'%s' % outputFilePath, mode='a', if_sheet_exists='overlay')
    # 操作第三张表
    # 1. 把RT HT复制到sheet3中，跳过1-7行和第13行空行。
    content1 = pd.read_excel(outputFilePath, sheet_name=excelSheetNameCPRT, skiprows=[0, 1, 2, 3, 4, 5, 6, 12],
                             header=None)
    content2 = pd.read_excel(outputFilePath, sheet_name=excelSheetNameCPHT, skiprows=[0, 1, 2, 3, 4, 5, 6, 12],
                             header=None)
    targetList1 = list(content1.loc[1])
    targetList2 = list(content2.loc[1])
    # 2. 循环删除包含VD_G2的列
    for num, i in enumerate(targetList1):
        if not (pd.isna(i)) and "VD_G2" in i:
            print("删除第一张表的列是:" + i + ",序号是:" + str(num))
            print()
            # inplace 会改变当前的dataframe,类似于给他重新赋值，并且返回none
            content1.drop(columns=num, inplace=True)
    for num, i in enumerate(targetList2):
        if not (pd.isna(i)) and "VD_G2" in i:
            print("删除第二张表的列是:" + i + ",序号是:" + str(num))
            # inplace 会改变当前的dataframe,类似于给他重新赋值，并且返回none
            content2.drop(columns=num, inplace=True)
    content1.to_excel(result, sheet_name="RE0505#02_COMP", index=False, header=None)
    content2.to_excel(result, sheet_name="RE0505#02_COMP", index=False, header=None, startcol=content1.shape[1],
                      startrow=0)
    print(content1)
    print(content2)
    result.save()
    print("写入成功")
