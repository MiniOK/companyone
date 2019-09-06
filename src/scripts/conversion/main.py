import os
import random
from shutil import copyfile

from win32com.client import Dispatch

from conversion.UltraSoundReport import UltraSound


def run_main():

    # DIR = "../../../data/20190722/"
    # DIR = "M:\\PycharmProjects\\doc_conversion\\data\\20190722\\"
    DIR = 'E:\\zjc\\'
    # Get all files in list
    word_documents = []
    excel_spreadsheets = []
    pdfs = []
    finish_file = []

    for version in os.listdir(DIR):
        if version == "V期":
            continue

        if "/" in DIR:
            sub_dir = "{}{}/".format(DIR, version)
        else:
            sub_dir = "{}{}\\".format(DIR, version)
        for category in os.listdir(sub_dir):
            if category == "Long_fu_survey_Carotid_ultrasound":
                continue
            print(category)
            if "/" in DIR:
                ssub_dir = sub_dir + category + "/"
            else:
                ssub_dir = sub_dir + category + "\\"
                print("ssub_dir",ssub_dir)
            for item in os.listdir(ssub_dir):
                if item == "2201":
                    print("继续执行！！！")

                    print("item",item)
                    if "/" in DIR:
                        sssub_dir = ssub_dir + item + "/"
                    else:
                        sssub_dir = ssub_dir + item + "\\"
                    for file in os.listdir(sssub_dir):
                        if "~$" not in file:
                            if file.split(".")[-1].lower() in ["doc", "docx"]:
                                if file.split(".")[-2] in ["doc", "docx"]:
                                    word_documents.append(sssub_dir + file)
                                    continue
                                word_documents.append(sssub_dir + file)
                            elif file.split(".")[-1].lower() in ["xls", "xlsx"]:
                                excel_spreadsheets.append(sssub_dir + file)
                            elif file.split(".")[-1].lower() in ["pdf"]:
                                pdfs.append(sssub_dir + file)
    # 保证每次执行都一致
    # random.seed(11)
    # 用于将一个列表中的元素打乱，即将列表中的元素随机排序
    # random.shuffle(word_documents)
    word = Dispatch("Word.Application")
    print("IV期高危的总数：", len(word_documents))



     # 将完成的文件读取
    for line in open("../output/finish.txt", encoding='utf-8'):
        finish_file.append(line)

    if len(finish_file) != len(set(finish_file)):
        print("重复前:", len(finish_file))
        print("有重复")
        finish_file = list(set(finish_file))
        print("去重后:",len(finish_file))
    else:
        print("没重复:", len(finish_file))


    # 过滤err中文件
    err_list = os.listdir('../output/err')


    try:
        finish_txt =  open('../output/finish-001.txt', 'r+', encoding='utf-8')
        data = finish_txt.read()
        # print(data)
        finish_txt.write(
            "{}\t\t{}\t\t\t{}\t\t\t{}\t\t{}\t\t\t{}\t\t\t{}\t\t\t{}\t\t\t{}\t\t\t{}\t\t\t{}\t\t\t{}\t\t\t{}\t\t\t{}\t\t\t{}\t\t\t{}\t\t\t"
            "{}\t\t\t{}\t\t\t{}{}\t\t\t{}\t\t\t{}\t\t\t{}\t\t\t{}\t\t\t{}\t\t\t{}\t\t\t{}\t\t\t{}\t\t\t{}\t\t\t{}\t\t\t{}\t\t\t{}\t\t\t{}\t\t\t{}\n"
            .format("name", "ID", "sex", "date", "v5", "v6", "v7", "v8", "v9", "v10", "v11", "v12", "v13", "v14","v15", "v16",
                    "v17","v18","v19","v20", "v21", "v22", "v23", "v24", "v25", "v26", "v27", "v28", "v29", "v30", "v31", "v32", "v33", "v34"))


        # 创建finish文件
        with open('../output/finish.txt', 'r+', encoding='utf-8') as finish:
            finish.read()
            for file in word_documents:
                finish_file_split = "{}\n".format(file.split("\\")[-1])
                # 过滤err中的文件
                if finish_file_split in err_list:
                    print("---------------过滤错误的文件")
                    continue
                # 过滤 完成文件
                if finish_file_split in finish_file:
                    # print("--------------过滤完成的文件")
                    continue
                #
                # if file == "E:\\zjc\\IV期\\Long_fu_survey_Carotid_ultrasound\\2202\\王欣华G2202301132.doc":
                #     continue
                # if file == "E:\\zjc\\IV期\\Long_fu_survey_Carotid_ultrasound\\2202\\郝守方G2202301624.doc":
                #     continue
                # if file == "E:\\zjc\\IV期\\Long_fu_survey_Carotid_ultrasound\\2202\\王志伟G2202300829.doc":
                #     continue
                # # 没有 RCCA-IMT 字段
                # if file == r"E:\zjc\IV期\High_risk_Carotid_ultrasound\1402\G1402404851报告单.docx":
                #     continue


                print(file)
                obj = UltraSound(file, word, finish, finish_txt)

    finally:
        finish.close()
        finish_txt.close()


if __name__ == '__main__':
    run_main()
    pass
