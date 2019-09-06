# -*- coding:utf-8 -*-
import os
import re
import warnings
from shutil import copyfile
import shutil
import xlrd
from win32com.client import Dispatch


class UltraSound:

    def __init__(self, file_path, word, finish, finish_txt):
        self.file_path = file_path
        self.title = "颈动脉超声检查报告"
        self.name = "Unknown"
        self.ID = "Unknown"
        self.gender = "Unknown"
        self.date = "Unknown"
        self.left_lcation = "左侧"
        self.CCA_IMT = "CCA-IMT(mm)"
        self.left_flag = "斑块（单位mm，空缺为正常）"
        self.CCA_IMT_left = "Normal"
        self.v1 = "NaN",
        self.v2 = "NaN",
        self.v3 = "NaN",
        self.plaques_count_left = "数量（1=单发，2=多发）"
        self.largest_plaque_width_left = "最大者长度"
        self.largest_plaque_depth_left = "最大者厚度"
        self.plaque_shape_left = "形态（1=规则型，2=不规则型）"
        self.plaque_is_ulcer_left = "是否溃疡型（0=否，1=是）"
        self.plaque_texture_left = "质地（A1均质低回声，A2均质等回声，A3=均质强回声，B=不均质"
        self.DS_left = "管腔直径狭窄率%"
        self.location_left = "狭窄部位"

        self.CCA_IMT_right = "Normal"
        self.v4 = "NaN",
        self.v5 = "NaN",
        self.v6 = "NaN",
        self.plaques_count_right = "数量（1=单发，2=多发）"
        self.largest_plaque_width_right = "最大者长度"
        self.largest_plaque_depth_right = "最大者厚度"
        self.plaque_shape_right = "形态（1=规则型，2=不规则型）"
        self.plaque_is_ulcer_right = "是否溃疡型（0=否，1=是）"
        self.plaque_texture_right = "质地（A1均质低回声，A2均质等回声，A3=均质强回声，B=不均质"
        self.DS_right = "管腔直径狭窄率%"
        self.location_right = "狭窄部位"

        self.comments = ""
        self.doctor = "Unknown"

        self.valid = True
        extension = file_path.split(".")[-1]
        if extension.lower() in ["doc", "docx"]:
            self.file_type = "word"
            self.load_doc(word, finish, finish_txt)
        elif extension.lower() in ["xls", "xlsx"]:
            self.file_type = "excel"
            self.load_xls()
        elif extension.lower() in ["pdf"]:
            self.file_type = "pdf"
        else:
            warn_msg = "Unknown file type with extension {} from {}. Please double check.".format(
                extension, file_path)
            warnings.warn(warn_msg)
            self.valid = False

    @staticmethod
    def fill(para, value):
        if value is None or value == "":
            return para
        else:
            return value

    def load_doc(self, word, finish, finish_txt):
        # word = Dispatch("Word.Application.8")
        word.Visible = 0
        try:
            f = word.Documents.Open(self.file_path)
        except Exception:
            copyfile(self.file_path,
                     "../output/err/{}".format(self.file_path.split("\\")[-1]))
            return 0
        content = f.Content.Text.replace("\t", "").replace("＝", "=").replace("\r", "").replace("\xa0", "").replace(
            "\x07", "").replace("官腔", "管腔").replace("\u3000", "").replace("\x00", "").replace("\x01",
                                                                                              "").replace(
            "\x15", "").replace("\x0c", "").replace("\x0e", "").replace("\x0c", "").replace("\x0b", "").replace(" ",
                                                                                                                "").replace(
            ":", "").replace("：", "").replace("端", "段").replace("%", "").replace("_", "").replace("。", "").replace("％","")
        if content[:1] != "广西医科大学附属武鸣医院":
            print("content", content)


            if len(content) < 10:
                copyfile(self.file_path,
                         "../output/err/{}".format(self.file_path.split("\\")[-1]))
                return 0
            # self.title = self.fill(self.title, re.search("南宁市武鸣区人民医院(.*?)姓名：", content).group(1).replace("\r", ""))
            try:
                self.name = self.fill(
                    self.name,
                    re.search(
                        "姓名(.*?)性别",
                        content).group(1).replace(
                        " ",
                        ""))
                if len(self.name) > 5:
                    raise ValueError
            except Exception:
                self.name = self.fill(
                    self.name,
                    re.search(
                        "姓名(.*?)(受检者ID|受检查ID)",
                        content).group(1).replace(
                        " ",
                        ""))
            print(self.name)
            try:
                self.ID = self.fill(
                    self.ID,
                    re.search(
                        "患者ID(.*?)左侧CCA-IMT",
                        content).group(1))
            except Exception:
                try:
                    self.ID = self.fill(
                        self.ID, re.search(
                            "受检者id|受检者ID(.*?)性别", content).group(1))
                except Exception:
                    copyfile(
                        self.file_path, "../output/err/{}".format(self.file_path.split("\\")[-1]))
                    return 0
            print(self.ID)
            try:
                self.gender = self.fill(
                    self.gender, re.search(
                        "性别(.*?)年龄", content).group(1))
            except Exception:
                self.gender = self.fill(
                    self.gender, re.search(
                        "性别(.*?)检查日期", content).group(1))
            print(self.gender)
            try:
                self.date = self.fill(
                    self.date, re.search(
                        "检查日期(.*?)打印日期", content).group(1))
            except Exception:
                self.date = self.fill(
                    self.date,
                    re.search(
                        "检查日期(.*?)(左侧|右侧|LCCA-IMT|2D及M型)",
                        content).group(1))
            self.date = self.date.replace(".", "-")
            print(self.date)





            if "LCCA-IMT" in content:
                if content.index("LCCA-IMT") < content.index("RCCA-IMT"):
                    left = re.search("LCCA-IMT(.*?)RCCA-IMT", content).group(1)
                    right = re.search("RCCA-IMT(.*?)超声印象", content).group(1)
                else:
                    right = re.search(
                        "RCCA-IMT(.*?)LCCA-IMT", content).group(1)
                    left = re.search("LCCA-IMT(.*?)超声印象", content).group(1)
            else:
                try:
                    if content.index("左侧") < content.index("右侧"):
                        left = re.search("左侧(.*?)右侧", content).group(1)

                        right = re.search(
                            "右侧(.*?)(超声印象|检查医生|报告医生)",
                            content).group(1)     # 添加 报告医生
                    else:
                        right = re.search("右侧(.*?)左侧", content).group(1)
                        left = re.search(
                            "左侧(.*?)(超声印象|检查提示)", content).group(1)
                except ValueError:
                    copyfile(
                        self.file_path, "../output/err/{}".format(self.file_path.split("\\")[-1]))
                    return 0


            print("left", left)
            try:
                self.CCA_IMT_left = (self.fill(self.CCA_IMT_left[0], re.search("近段(.*?)中段", left).group(1)),
                                     self.fill(
                    self.CCA_IMT_left[1], re.search(
                        "中段(.*?)远段", left).group(1)),
                    self.fill(self.CCA_IMT_left[2], re.search("远段(.*?)(数量|斑块)", left).group(1)))

            except Exception:
                try:
                    self.CCA_IMT_left = (self.fill(self.CCA_IMT_left[0], re.search("近段(.*?)mm", left).group(1)),
                                         self.fill(
                        self.CCA_IMT_left[1], re.search(
                            "中段(.*?)mm", left).group(1)),
                        self.fill(self.CCA_IMT_left[2], re.search("远段(.*?)mm", left).group(1)))
                except Exception:
                    copyfile(
                        self.file_path, "../output/err/{}".format(self.file_path.split("\\")[-1]))
                    return 0
            self.v1 = self.CCA_IMT_left[0]
            self.v2 = self.CCA_IMT_left[1]
            self.v3 = self.CCA_IMT_left[2]
            print(self.CCA_IMT_left)

            # 数量
            try:

                self.plaques_count_left = self.fill(self.plaque_is_ulcer_left,
                                                    re.search("(数量1.无2.单发3.多发|数量（1=单发，2=多发）|数量（1＝单发，2＝多发）)(.*?)(最大者长度|最长者长度|最大者厚度)", left).group(2))
            except Exception:
                self.plaques_count_left = self.fill(self.plaque_is_ulcer_left,
                                                    re.search("(数量1.无2.单发3.多发|数量（1=单发，2=多发）|数量（1＝单发，2＝多发）)(.*?)(最大者长度|最长者长度|最大者厚度)", left).group(2))
            print(self.plaques_count_left)



            # 最大者长度
            try:
                self.largest_plaque_width_left = self.fill(self.largest_plaque_width_left,
                                                           re.search("(最大者长度|大者长度|最大者长|最长者长度)(.*?[0-9]+.[0-9]+)(mm|最大者厚度|最大厚度|大者厚度)", left).group(2))
            except Exception:
                self.largest_plaque_width_left = self.fill(self.largest_plaque_width_left,
                                                           re.search("(最大者长度|大者长度|最大者长|最长者长度)(.*?)(mm|最大者厚度|最大厚度|大者厚度)", left).group(2))
            print(self.largest_plaque_width_left)



            # 最大者厚度
            try:
                self.largest_plaque_depth_left = self.fill(self.largest_plaque_depth_left,
                                                           re.search("(最大者厚度|大者厚度|最大厚度)(.*?)(mm。|回声|形态)", left).group(2))
            except Exception:
                self.largest_plaque_depth_left = self.fill(self.largest_plaque_depth_left,
                                                           re.search("(最大者厚度|大者厚度|最大厚度)(.*?)(mm。|回声|形态)", left).group(2))
            print(self.largest_plaque_depth_left)




            # 形态
            try:
                self.plaque_shape_left = self.fill(self.plaque_shape_left,
                                                   re.search("(形态（1=规则型，2=不规则型）|形态1.规则型2.不规则型|形态（1=规则型，2=不规则型|（1=规则型，2=不规则）)(.*?)(是1否溃疡型|有无溃疡斑块|形态（1=规则型，2=不规则型。）|是否溃疡型)",
                                                        left).group(2))
            except Exception:
                self.plaque_shape_left = self.fill(self.plaque_shape_left,
                                                   re.search(
                                                       "(形态（1=规则型，2=不规则型）|形态1.规则型2.不规则型|形态（1=规则型，2=不规则型|（1=规则型，2=不规则）)(.*?)(有无溃疡斑块|形态（1=规则型，2=不规则型。）|是否溃疡型)",
                                                       left).group(2))
            print(self.plaque_shape_left)




            # 是否溃疡
            try:
                self.plaque_is_ulcer_left = self.fill(self.plaque_is_ulcer_left,
                                                      re.search("(是否溃疡型（=否，1=是）|是1否溃疡型（0=否，1=是）|是否溃疡型（0=否，1=是）|有无溃疡斑块1.无2.有|是否溃疡型（0＝否，1＝是）)(.*?)(A1|狭窄程度|质地)", left).group(2))
            except Exception:
                self.plaque_is_ulcer_left = self.fill(self.plaque_is_ulcer_left,
                                                      re.search(
                                                          "(是否溃疡型（0=否，1=是）|有无溃疡斑块1.无2.有|是否溃疡型（0＝否，1＝是）)(.*?)(狭窄程度|质地)",
                                                          left).group(2))
            print(self.plaque_is_ulcer_left)



            # 质地
            try:
                self.plaque_texture_left = self.fill(self.plaque_texture_left,
                                                     re.search("(质地（A1=均质低回声A2=均质等回声，A3=均质强回，声，B=不均质）|质地（A1=均质低回声，A2=均质等回声，A3=均质强回声，B=不均质A3|质地（A1=均质低回声，A2=均质等回声，A3=均质强回声2，B=不均质）|1.强回声2.中等回声3.低回声4.不均匀回声|质地（A1=均质低回声，A2=均质等1回声，A3=均质强回声，B=不均质）|质地（A1均质低回声，A2均质等回声，A3=均质强回声，B=不均质）|质地（A1=均质低回声，A2=均质等回声，A3=均质强回声，B=不均匀）|质地（A1=均质低回声，A2=均质等回声，A3=均质强回声，B=不均质）|质地（A1=均质低回声，A2=均质等回声，A3均质强回声，B=不均质）)(.*?)(管腔直径狭窄率%|形态1.规则型|官腔直径狭窄率|管腔直径狭窄率|。管腔直径狭窄率)",
                                                               left).group(2))
            except Exception:
                self.plaque_texture_left = self.fill(self.plaque_texture_left,
                                                     re.search(
                                                         "(质地（A1=均质低回声，A2=均质等回声，A3=均质强回声，B=不均质A3|质地（A1=均质低回声，A2=均质等回声，A3=均质强回声2，B=不均质）|1.强回声2.中等回声3.低回声4.不均匀回声|质地（A1=均质低回声，A2=均质等1回声，A3=均质强回声，B=不均质）|质地（A1均质低回声，A2均质等回声，A3=均质强回声，B=不均质）|质地（A1=均质低回声，A2=均质等回声，A3=均质强回声，B=不均匀）|质地（A1=均质低回声，A2=均质等回声，A3=均质强回声，B=不均质）|质地（A1=均质低回声，A2=均质等回声，A3均质强回声，B=不均质）)(.*?)(管腔直径狭窄率%|形态1.规则型|官腔直径狭窄率|管腔直径狭窄率|。管腔直径狭窄率)",
                                                         left).group(2))
            print(self.plaque_texture_left)

            # 官腔直径狭窄率
            try:
                self.DS_left = self.fill(
                    self.DS_left, re.search(
                        "(管腔直径狭窄率|狭窄程度或闭塞部位|官腔直径狭窄率|管腔直径狭窄率%|管腔直径狭窄率％)(.*?)(狭窄部位|检查结果)", left).group(2))
            except Exception:
                self.DS_left = self.fill(
                    self.DS_left, re.search(
                        "(管腔直径狭窄率|狭窄程度或闭塞部位|官腔直径狭窄率|管腔直径狭窄率%)(.*?)(狭窄部位|检查结果)", left).group(2))
            print(self.DS_left)




            try:
                self.location_left = self.fill(
                    self.location_left, re.search(

                         "(狭窄部位|狭窄程度或闭塞部位)(.*?)", left).group(2))
            except Exception:
                self.location_left = self.fill(
                    self.location_left, re.search(
                        "狭窄部位|狭窄程度或闭塞部位(.*?)", left).group(2))
            print(self.location_left)



            for t in f.Tables:
                try:
                    self.doctor = self.fill(self.doctor, t.Cell(20, 6).Range.Text)
                    # print("1", self.doctor)
                except Exception:
                    # self.doctor = re.search(content, "报告医生：(.*?)").group(1)
                    try:
                        self.doctor = re.search("(报告医生:|报告医生：|报告医师:|报告医师)(.*?)\r", f.Content.Text).group(2)
                        # print("2", self.doctor)
                    except Exception:
                        try:
                            self.doctor = re.search("(报告医生|报告医生 |报告医师 |检查医生)(.*?)\r", f.Content.Text).group(2)
                            # print("3", self.doctor)
                            # 添加报告医生
                        except Exception:
                            try:
                                self.doctor = re.search("\r\x07\r\x07\r(报告医生：|报告医师：)(.*?)报告机构", f.Content.Text).group(
                                    1).replace(" ", "")
                                # print("4", self.doctor)
                            except Exception:
                                try:
                                    self.doctor = re.search("(报告医生：|报告医师：)(.*?)\r", f.Content.Text).group(
                                        1).replace(" ", "")
                                except Exception:
                                    copyfile(self.file_path,
                                             "../../../output/errr/{}".format(self.file_path.split("\\")[-1]))
                                    print("doctor", self.file_path)
                break

            self.doctor = self.doctor.replace(" ", "").replace("\n", "").replace("\r", "").replace("：","")
            print(self.doctor)

            # 将完成的文件逐行写入finish
            finish_file = '{}\n'.format(self.file_path.split("\\")[-1])
            finish.write(finish_file)

            finish_txt.write(
                "{}\t\t{}\t\t{}\t\t{}\t\t{}\t\t{}\t\t{}\t\t{}\t\t{}\t\t{}\t\t{}\t\t{}\t\t{}\t\t{}\t\t{}\t\t{}\t\t{}\t\t{}\t\t{}\n"
                    .format(self.name,
                            self.ID,
                            self.gender,
                            self.date,
                            self.left_lcation,
                            self.CCA_IMT,
                            # self.CCA_IMT_left,
                            self.v1,
                            self.v2,
                            self.v3,
                            self.left_flag,
                            self.plaques_count_left,
                            self.largest_plaque_width_left,
                            self.largest_plaque_depth_left,
                            self.plaque_shape_left,
                            self.plaque_is_ulcer_left,
                            self.plaque_texture_left,
                            self.DS_left,
                            self.location_left,
                            self.doctor
                            ))


            f.Close()

if __name__ == "__main__":
    word = Dispatch("Word.Application")
    test = UltraSound(

        r"E:\zjc\IV期\High_risk_Carotid_ultrasound\1102\11020600023.doc",
        word, finish=None, finish_txt=None)
































        # print("right", right)
            # try:
            #     self.CCA_IMT_right = (self.fill(self.CCA_IMT_right[0], re.search("近段(.*?)中段", right).group(1)),
            #                           self.fill(self.CCA_IMT_right[1], re.search("中段(.*?)远段", right).group(1)),
            #                           self.fill(self.CCA_IMT_right[2], re.search("远段(.*?)(数量|斑块)", right).group(1)))
            # except Exception as e:
            #     try:
            #         self.CCA_IMT_right = (self.fill(self.CCA_IMT_right[0], re.search("近段(.*?)mm", right).group(1)),
            #                               self.fill(self.CCA_IMT_right[1], re.search("中段(.*?)mm", right).group(1)),
            #                               self.fill(self.CCA_IMT_right[2], re.search("远段(.*?)mm", right).group(1)))
            #     except Exception:
            #         copyfile(self.file_path, "../output/err/{}".format(self.file_path.split("\\")[-1]))
            #         return 0
            # # print(self.CCA_IMT_right)
            # try:
            #     self.plaques_count_right = self.fill(self.plaques_count_right,
            #                                          re.search("数量（1=单发，2=多发）[(.*?)]", right).group(1))
            # except Exception:
            #     try:
            #         self.plaques_count_right = self.fill(self.plaques_count_right,
            #                                              re.search("数量1.无2.单发3.多发(.*?)(最大者长度|最长者长度)", right).group(1))
            #     except Exception:
            #         self.plaques_count_right = self.fill(self.plaques_count_right,
            #                                              re.search("数量（1=单发，2=多发）(.*?)(最大者长度|最长者长度)", right).group(1))
            # # print(self.plaques_count_right)
            # try:
            #     self.largest_plaque_width_right = self.fill(self.largest_plaque_width_right,
            #                                                 re.search("(最大者长度|最长者长度)(.*?)mm", right).group(1))
            # except Exception:
            #     try:
            #         self.largest_plaque_width_right = self.fill(self.largest_plaque_width_right,
            #                                                     re.search("(最大者长度|最长者长度)(.*?)(最大者厚度|最大厚度|最大着厚度)",
            #                                                               right).group(1))
            #     except Exception:
            #         raise Exception
            # # print(self.largest_plaque_width_right)
            # try:
            #     self.largest_plaque_depth_right = self.fill(self.largest_plaque_depth_right,
            #                                                 re.search("最大者厚度(.*?)mm。", right).group(1))
            # except Exception:
            #     try:
            #         self.largest_plaque_depth_right = self.fill(self.largest_plaque_depth_right,
            #                                                     re.search("最大者厚度(.*?)回声", right).group(1))
            #     except Exception:
            #         # try:
            #         regex = re.compile("(最大着厚度|最大者厚度|最大厚度)(.*?)形态")
            #         self.largest_plaque_depth_left = self.fill(self.largest_plaque_depth_left,
            #                                                    regex.search(right).group(1))
            #             # 修改 NoneType 异常
            #         # except Exception:
            #         #     warn_msg = "Warnings come form {} ".format(self.file_path)
            #         #
            #         #     warnings.warn(warn_msg)
            #         #     import time
            #         #
            #         #     time.sleep(60)
            #         #     print('正在等待。。。')
            # # print(self.largest_plaque_depth_right)
            # try:
            #     self.plaque_shape_right = self.fill(self.plaque_shape_right,
            #                                         re.search("形态（1=规则型，2=不规则型）[(.*?)]", right).group(1))
            # except Exception:
            #     try:
            #         self.plaque_shape_right = self.fill(self.plaque_shape_right,
            #                                             re.search("形态1.规则型2.不规则型(.*?)有无溃疡斑块", right).group(1))
            #     except Exception:
            #         self.plaque_shape_right = self.fill(self.plaque_shape_right,
            #                                             re.search(
            #                                                 "(形态（1=规则型，2=不规则型|形态（1=规则型，2=不规则型）|形态（1=规则型，2=不规则）|形态（1=规则型，2=不规则型。）)(.*?)是否溃疡",
            #                                                 right).group(1))
            # # print(self.plaque_shape_right)
            # try:
            #     self.plaque_is_ulcer_right = self.fill(self.plaque_is_ulcer_right,
            #                                            re.search("是否溃疡型（0=否，1=是）[(.*?)]", right).group(1))
            # except Exception:
            #     try:
            #         self.plaque_is_ulcer_right = self.fill(self.plaque_is_ulcer_right,
            #                                                re.search("有无溃疡斑块1.无2.有(.*?)狭窄程度", right).group(1))
            #     except Exception:
            #         self.plaque_is_ulcer_right = self.fill(self.plaque_is_ulcer_right,
            #                                                re.search("(是否溃疡型（0=否，1=是）|是否溃疡（0=否，1=是）)(.*?)质地",
            #                                                          right).group(1))
            # # print(self.plaque_is_ulcer_right)
            # try:
            #     self.plaque_texture_right = self.fill(self.plaque_texture_right,
            #                                           re.search("质地（A1=均质低回声，A2=均质等回声，A3均质强回声，B=不均质）[(.*?)]",
            #                                                     right).group(1))
            # except Exception:
            #     try:
            #         self.plaque_texture_right = self.fill(self.plaque_texture_right,
            #                                               re.search("1.强回声2.中等回声3.低回声4.不均匀回声(.*?)形态1.规则型", right).group(
            #                                                   1))
            #     except Exception:
            #         self.plaque_texture_right = self.fill(self.plaque_texture_right,
            #                                               re.search(
            #                                                   "(质地（A1=均质低回声，A2=均质等回声，A3=均质强回声，B=不均匀）|质地（A1均质低回声，A2均质等回声，A3=均质强回声，B=不均质）|质地（A1=均质低回声，A2=均质等回声，A3=均质强回声，B=不均质）|质地（A1=均质低回声，A2=均质等回声，A3均质强回声，B=不均质）\[)(.*?)(管腔直径狭窄率|]。管腔直径狭窄率|狭窄部位)",
            #                                                   right).group(
            #                                                   1))
            #
            # # print(self.plaque_texture_right)
            # try:
            #     self.DS_right = self.fill(self.DS_right, re.search("管腔直径狭窄率(.*?)(%狭窄部位|狭窄部位)", right).group(1))
            # except Exception:
            #     try:
            #         self.DS_right = self.fill(self.DS_right, re.search("狭窄程度或闭塞部位(.*?)(检查结果|左侧)", right).group(1))
            #     except Exception:
            #         try:
            #             self.DS_right = self.fill(self.DS_right, re.search("管腔直径狭窄率(.*?)狭窄部位", right).group(1))
            #         except Exception:
            #             self.DS_right = self.fill(self.DS_right, re.search("(狭窄程度或闭塞部位|狭窄部位)(.*?)", right).group(1))
            # # print(self.DS_right)
            # try:
            #     self.location_right = self.fill(self.location_right, re.search("狭窄部位(.*?)", right).group(1))
            # except Exception:
            #     try:
            #         self.location_right = self.fill(self.location_right,
            #                                         re.search("狭窄程度或闭塞部位(.*?)检查结果", right).group(1))
            #     except Exception:
            #         self.location_right = self.fill(self.location_right,
            #                                         re.search("狭窄程度或闭塞部位(.*?)左侧", content).group(1))
            # # print(self.location_right)
            #
            # try:
            #     self.comments = self.fill(self.comments, re.search("狭窄部位|超声印象(.*?)", content).group(1)) # 添加 空串 ''
            # except Exception as e:
            #     print(e)
            #     self.comments = self.fill(self.comments, re.search("检查提示(.*?)检查医生", content).group(1))
            #
