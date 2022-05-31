import pandas as pd
import os

# AE7252#05-CPRT AE7252#05-CPHT
desktopPath = os.path.join(os.path.expanduser('~'), "Desktop")
inputFilePath = desktopPath + r'\MERGE'
outputFilePath = desktopPath + r'\AE7252#05-COMP1.xlsx'
if os.path.exists(r'%s' % outputFilePath):
    os.remove(r'%s' % outputFilePath)
# 这个地方创建了一个目标文件（说实话我也不是很清楚为什么这么写）
result = pd.ExcelWriter(r'%s' % outputFilePath)
origin_file_list = os.listdir(r'%s' % inputFilePath)
# 这个地方暴力处理了一下，直接反转，因为需求是RT排在前面HT排在后面，我在考虑python是否可以像java一样自己重写排序方法
origin_file_list.reverse()
for i in origin_file_list:
    excel_file_name = i.replace("-", "_").replace(".xlsx", "")
    file_path = r'%s' % inputFilePath + "\\" + i
    # header=None代表不处理header
    content = pd.read_csv(file_path, header=None,skip_blank_lines=False)
    print(content.iloc[[1],[3]])
    # index=False代表不添加索引
    # content.to_excel(result, excel_file_name, index=False, header=None)
# df = pd.DataFrame()
# print("请输入第三个sheet的名字：")
# sheet3 = input()
# df.to_excel(result, sheet3)
# df.to_excel(result, "RT_fail")
# df.to_excel(result, "HT_fail")
# df.to_excel(result, "良品数量")
# result.save()
