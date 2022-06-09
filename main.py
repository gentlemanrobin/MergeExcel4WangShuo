import pandas as pd
import os

desktopPath = os.path.join(os.path.expanduser('~'), "Desktop")
outputFilePath = desktopPath + r'\AE7252#05-COMP1.xlsx'
writer = pd.ExcelWriter(r'%s' % outputFilePath, mode='a', if_sheet_exists='overlay')
content1 = pd.read_excel(outputFilePath, sheet_name=0, skiprows=[0, 1, 2, 3, 4, 5, 6, 12], header=None)
content2 = pd.read_excel(outputFilePath, sheet_name=1, skiprows=[0, 1, 2, 3, 4, 5, 6, 12], header=None)
targetList1 = list(content1.loc[1])
targetList2 = list(content2.loc[1])
# 循环删除多列
for num, i in enumerate(targetList1):
    if not (pd.isna(i)) and "VD_G2" in i:
        print("删除第一张表的列是:" + i+",序号是:" + str(num))
        print()
        # inplace 会改变当前的dataframe,类似于给他重新赋值，并且返回none
        content1.drop(columns=num, inplace=True)
for num, i in enumerate(targetList2):
    if not (pd.isna(i)) and "VD_G2" in i:
        print("删除第二张表的列是:" + i+",序号是:" + str(num))
        # inplace 会改变当前的dataframe,类似于给他重新赋值，并且返回none
        content2.drop(columns=num, inplace=True)
content1.to_excel(writer, sheet_name="RE0505#02_COMP", index=False, header=None)
content2.to_excel(writer, sheet_name="RE0505#02_COMP", index=False, header=None, startcol=content1.shape[1], startrow=0)


# print(content1.loc[:, 1])
writer.save()

# lists = ["nan", "nan", 'SITE(15)', 'X(15)', 'Y(15)', 'Gate_CONT(11)', 'Source_CONT(11)', 'DRAIN_CONT(11)',
#          'SUB_CONT(11)', 'GA_CONT(11)', 'IGA(12)', 'VD_G1(14)', 'IGOFF1(15)', 'IGON1(16)', 'VTH1(17)', 'RDON1(18)',
#          'ID100V1(19)', 'IS100V1(19)', 'IG100V1(19)', 'ISUB100V1(19)', 'ID200V1(20)', 'IS200V1(20)', 'IG200V1(20)',
#          'ISUB200V1(20)', 'ID300V1(21)', 'IS300V1(21)', 'IG300V1(21)', 'ISUB300V1(21)', 'ID400V1(22)', 'IS400V1(22)',
#          'IG400V1(22)', 'ISUB400V1(22)', 'ID500V1(23)', 'IS500V1(23)', 'IG500V1(23)', 'ISUB500V1(23)', 'ID600V1(24)',
#          'IS600V1(24)', 'IG600V1(24)', 'ISUB600V1(24)', 'ID700V1(25)', 'IS700V1(25)', 'IG700V1(25)', 'ISUB700V1(25)',
#          'ID800V1(26)', 'IS800V1(26)', 'IG800V1(26)', 'ISUB800V1(26)', 'VTH2(27)', 'RDON2(28)', 'ID750V2(29)',
#          'RDON3(30)', 'VD_G2(32)', 'IGON2(33)', 'IGOFF2(34)', 'ID100V2(35)', 'IS100V2(35)', 'IG100V2(35)',
#          'ISUB100V2(35)', 'VTH2-1(36)', 'R2/R1(37)', 'R3/R1(38)']
# for num, i in enumerate(lists):
#     if "VD_G2" in i:
#         print("删除的列是:" + i)
#         print("序号是:" + str(num))
# lists1 = [1,2,3,4,5]
# for num in lists1:
#     print(num)
