import pandas as pd
import os


# 打包命令：因为会用到'C:\WINDOWS\system32\api-ms-win-crt-locale-l1-1-0.dll文件，所以需要进到C:\WINDOWS\system32目录打包
# pyinstaller -F -w C:\Users\RobinGao\Desktop\Code\python\first\mergeExcel.py
# pyinstaller -F C:\Users\RobinGao\Desktop\Code\python\first\mergeExcel.py --distpath=C:\Users\RobinGao\Desktop\Code\python\first\dist


def get_sheet_name(sheetName):
    excelSheetName = sheetName.replace(".xlsx", "")
    print("准备处理：" + excelSheetName)
    # 下面对文件名做初步处理，用find查找是否包含#
    # 如果包含就往前取6个字符（批号）往后取2个字符（片号）,如果不包含会返回-1
    targetCharIndex = excelSheetName.find('#')
    if targetCharIndex != -1 and len(excelSheetName) >= 9:
        name1 = excelSheetName[targetCharIndex - 6:targetCharIndex]
        name2 = excelSheetName[targetCharIndex + 1:targetCharIndex + 3]
        excelSheetName = name1 + "#" + name2
        if 'CPH' in name:
            excelSheetName = excelSheetName + "_CPHT"
        elif 'CPR' in name:
            excelSheetName = excelSheetName + "_CPRT"
    else:
        print("文件名不符合规则")

    print("excelSheetName==" + excelSheetName)
    return excelSheetName


if __name__ == '__main__':
    desktopPath = os.path.join(os.path.expanduser('~'), "Desktop")
    inputFilePath = desktopPath + r'\MERGE'
    outputFilePath = desktopPath + r'\AE7252#05-COMP1.xlsx'
    if os.path.exists(r'%s' % outputFilePath):
        os.remove(r'%s' % outputFilePath)
    # 这个地方创建了一个目标文件（说实话我也不是很清楚为什么这么写）
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
        # 取单元格的值(下面两种方式都可以)
        # name = content.loc[[1], [3]].values[0][0]
        name = content.loc[1][3]
        print("单元格名字" + name)
        # CPRA1应该排在前面，所以如果是CPHA1出现了，此时要重新读取第二个文件（就是CPRA1了）
        if 'CPH' in name:
            excelSheetNameCPHT = get_sheet_name(i)
            contentCPHT = content
        elif 'CPR' in name:
            excelSheetNameCPRT = get_sheet_name(i)
            contentCPRT = content

    # index=False代表不添加索引
    contentCPRT.to_excel(result, excelSheetNameCPRT, index=False, header=None)
    contentCPHT.to_excel(result, excelSheetNameCPHT, index=False, header=None)
    df = pd.DataFrame()
    print("请输入第三个sheet的名字：")
    sheet3 = input()
    df.to_excel(result, sheet3)
    df.to_excel(result, "RT_fail")
    df.to_excel(result, "HT_fail")
    df.to_excel(result, "良品数量")
    result.save()
