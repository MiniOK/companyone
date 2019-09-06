import random
import re

def main():
    content = 'Hello, I am Jerry, from Chongqing, a montain city, nice to meet you……'
    regex = re.compile('\w*o\w*')
    z = regex.search(content)
    print(z)
    print(type(z))
    print(z.group())
    print(z.span())

import os

DIR = '../output/err'
err_list = os.listdir(DIR)
print(len(err_list))
for i in err_list:
    print(i)
    print(type(i))
print(type(err_list))



# if __name__ == '__main__':
#     string = "最大者长度12mm"
#
#     # regex = re.compile("最长者长度(.*?)最长者长度")
#     regex = re.compile("最大者长度(.*?)mm")

    # print(regex.search(string).group(1))


# list = ['A', 'B', 'C', 'D', 'E']
#
# random.seed(11)
# random.shuffle(list)
# print(list)


# def return_test(x):
#     if x > 0:
#         return x
#     else:
#         return 0
#
#
# if __name__ == '__main__':
#     print(return_test(-2))


'''
content 
颈动脉超声检查报告
姓名魏振平
受检者ID
G1508507939
性别男
检查日期2018-12-25
左侧
CCA-IMT(mm)
近段0.6中段0.7远段0.6
斑块（单位mm，空缺为正常）
数量（1=单发，2=多发）
最大者长度最大者厚度
形态（1=规则型，2=不规则型）
是否溃疡型（0=否，1=是）
质地（A1=均质低回声，A2=均质等回声，A3=均质强回声，B=不均质）
管腔直径狭窄率%
狭窄部位
右侧
CCA-IMT(mm)
近段0.5中段0.5远段0.5
斑块（单位mm，空缺为正常）
数量（1=单发，2=多发）
最大者长度最大者厚度
形态（1=规则型，2=不规则型）
是否溃疡型（0=否，1=是）
质地（A1=均质低回声，A2=均质等回声，A3=均质强回声，B=不均质）
管腔直径狭窄率%
狭窄部位 
超声印象双侧颈总动脉未见异常
报告医生赵蓓

'''
