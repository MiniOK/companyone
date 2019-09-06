# 文件复制
# -*- coding:utf-8 -*-
import shutil

# file_path = r"C:\Users\miniloveliness\Desktop\test_code\test.txt"
# new_path = r"C:\Users\miniloveliness\Desktop\new.txt"
# shutil.copyfileobj(open(file_path, 'r', encoding='utf-8'), open(new_path, 'w', encoding='utf-8'))


# class EvaException(BaseException):
# #     def __init__(self,msg):
# #         self.msg=msg
# #     def __str__(self):
# #         return self.msg
# #
# # try:
# #     raise EvaException('类型错误')
# # except EvaException as e:
# #     print(e)
if __name__ == '__main__':

    with open('../output/finish.txt', 'w', encoding='utf-8') as finish:
        print(finish)
        print('1111111111111111')
        finish.write('1111\n')
        finish.write('1111')


