import pandas as pd
import os


# 打包命令：因为会用到'C:\WINDOWS\system32\api-ms-win-crt-locale-l1-1-0.dll文件，所以需要进到C:\WINDOWS\system32目录打包
# pyinstaller -F -w C:\Users\RobinGao\Desktop\Code\python\first\mergeExcel.py
# pyinstaller -F C:\Users\RobinGao\Desktop\Code\python\first\mergeExcel.py --distpath=C:\Users\RobinGao\Desktop\Code\python\first\dist
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
    print("sheetName=="+sheetName)
    return sheetName


if __name__ == '__main__':
    desktopPath = os.path.join(os.path.expanduser('~'), "Desktop")
    inputFilePath = desktopPath + r'\MERGE'
    outputFilePath = desktopPath + r'\AE7252#05-COMP1.xlsx'
    if os.path.exists(r'%s' % outputFilePath):
        os.remove(r'%s' % outputFilePath)
    # 创建一个目标文件（说实话我也不是很清楚为什么这么写）
    result = pd.ExcelWriter(r'%s' % outputFilePath)
    origin_file_list = os.listdir(r'%s' % inputFilePath)
    # 定义4个临时存储文件内容变量,如果后期文件多，要考虑用冒泡排序，而不是手动to_excel
    contentCPRT = ""
    contentCPHT = ""
    excelSheetNameCPRT = ""
    excelSheetNameCPHT = ""

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
    df.to_excel(result, get_sheet_previous_name(origin_file_list[0])+"_COMP")
    df.to_excel(result, "RT_fail")
    df.to_excel(result, "HT_fail")
    df.to_excel(result, "良品数量")
    result.save()
