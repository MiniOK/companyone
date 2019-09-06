# coding: utf-8

# with open("../output/finish.txt", "wt", encoding="utf-8") as f:
#     # print() 函数的输出重定向到一个文件中 只需要在print函数上加入 file 关键字即可：
#     print("nihaonihaonihaonihaonihaonihaonihao", file=f)


for i in range(10):
    print(i, end=' ')
print(1,2,3,4,5, sep=",")

row = ("nihao", 'wohao',"ahoa")
print(",".join(row))


str1 = """content 颈动脉超声检查报告姓名
王先拽受检者IDG1404408211性别女检查日期2018-3-8
左侧CCA-IMT(mm)
近段0.9中段0.8远段1.0数量（1=单发，2=多发）
最大者长度最大者厚度形态（
1=规则型，2=不规则型）是否溃疡型（0=否，1=是）
质地（A1=均质低回声，A2=均质等回声，A3=均质强回声，
B=不均质）管腔直径狭窄率狭窄部位右侧CCAIMT(mm)近段0.9中段0.9远段1.1
右侧膨大处内膜厚约1.3mm数量（1=单发，2=多发）
最大者长度最大者厚度形态（1=规则型，2=不规则型）
是否溃疡型（0=否，1=是）质地（A1=均质低回声，
A2=均质等回声，A3=均质强回声，B=不均质）
管腔直径狭窄率狭窄部位超声印象右侧颈动脉膨
大处内膜局限性增厚仪器型号PHILIPS－HD9(2013年购置)报告医生何晋阳
"""
if "左侧CCA-IMT(mm)" in str1:
    print("That's OK!")