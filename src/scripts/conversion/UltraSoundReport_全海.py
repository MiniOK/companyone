# -*- coding:utf-8 -*-
import warnings
from shutil import copyfile
import re
from win32com.client import Dispatch


class UltraSound:
    def __init__(self, file_path, word):
        self.file_path = file_path
        self.title = "颈动脉超声检查报告"
        self.name = "找不到姓名"
        self.ID = "找不到受检者ID"
        self.gender = "找不到性别"
        self.date = "找不到检查日期"
        self.CCA_IMT_left = "Normal"
        self.plaques_count_left = "Normal"
        self.largest_plaque_width_left = "Normal"
        self.largest_plaque_depth_left = "Normal"
        self.plaque_shape_left = "Normal"
        self.plaque_is_ulcer_left = "Normal"
        self.plaque_texture_left = "Normal"
        self.DS_left = "Normal"
        self.location_left = "Normal"



        self.CCA_IMT_right = "Normal"
        self.plaques_count_right = "Normal"
        self.largest_plaque_width_right = "Normal"
        self.largest_plaque_depth_right = "Normal"
        self.plaque_shape_right = "Normal"
        self.plaque_is_ulcer_right = "Normal"
        self.plaque_texture_right = "Normal"
        self.DS_right = "Normal"
        self.location_right = "Normal"

        self.comments = ""
        self.doctor = "Unknown"

        self.valid = True
        extension = file_path.split(".")[-1]
        if extension.lower() in ["doc", "docx"]:
            self.file_type = "word"
            self.load_doc(word)
        elif extension.lower() in ["xls", "xlsx"]:
            self.file_type = "excel"
            # self.load_xls()
        elif extension.lower() in ["pdf"]:
            self.file_type = "pdf"
        else:
            warn_msg = "Unknown file type with extension {} from {}. Please double check.".format(extension, file_path)
            warnings.warn(warn_msg)
            self.valid = False

    @staticmethod
    def fill(para, value):
        if value is None or value == "":
            return para
        else:
            return value

    def load_doc(self, word):
        word.Visible = 0
        # 打开文档
        try:
            f = word.Documents.Open(self.file_path)
        except Exception:
            copyfile(self.file_path, "../../../output/err/open_fail/{}".format(self.file_path.split("\\")[-1]))
            with open("../../../output/err/err.txt", "w") as f:
                f.write("{} 文件无法打开 \n".format(self.file_path))
                return 0
        # 文档转换为字符串
        try:
            content = f.Content.Text.replace("\t", "").replace("＝", "=").replace("\r", "").replace("\xa0", "").replace(
                "\x07", "").replace("官腔", "管腔").replace("\u3000", "").replace("\x00", "").replace("\x01", "").replace(
                "\x15", "").replace("\x0c", "").replace("\x0e", "").replace("\x0c", "").replace("\x0b", "").replace(" ",
                                                                                                                    "").replace(
                ":", "").replace("：", "").replace("端", "段")
        except Exception:
            copyfile(self.file_path, "../../../output/errr/{}".format(self.file_path.split("\\")[-1]))
            return 0
        if content[:1] != "广西医科大学附属武鸣医院":
            # 处理空文档情况
            if len(content) < 10:
                copyfile(self.file_path, "../../../output/err/none_file/{}".format(self.file_path.split("\\")[-1]))
                with open("../../../output/err/err.txt", "a") as f:
                    f.write("{} {}\n".format(self.file_path, "文档空白"))
                return 0
            # print(content)
            # 提取 name
            try:
                self.name = self.fill(self.name,
                                      re.search("姓名(.*?)(受检者ID|受检查ID)", content).group(1).replace(" ", ""))
            except Exception:
                try:
                    self.name = self.fill(self.name,
                                          re.search("模板颈动脉超声检查报告(.*?)受检者ID", content).group(1))
                except Exception:
                    copyfile(self.file_path,
                             "../../../output/errr/{}".format(self.file_path.split("\\")[-1]))
                    print("name", self.file_path)
                    return 0
            #  提取 ID
            try:
                self.ID = self.fill(self.ID, re.search("患者ID(.*?)左侧CCA-IMT", content).group(1))
            except Exception:
                try:
                    self.ID = self.fill(self.ID, re.search("(受检者ID|受检查ID)(.*?)性别", content).group(2))
                except Exception:
                    copyfile(self.file_path,
                             "../../../output/errr/picture_no_string/{}".format(self.file_path.split("\\")[-1]))
                    with open("../../../output/err/err.txt", "a") as f:
                        f.write("{}{}\n".format(self.file_path, "图片非字符"))
                    print("ID", self.file_path)
            # 提取 gender
            try:
                self.gender = self.fill(self.gender, re.search("性别(.*?)年龄", content).group(1))
            except Exception:
                try:
                    self.gender = self.fill(self.gender, re.search("性别(.*?)(日期|检查日期|时间)", content).group(1))
                except Exception:
                    try:
                        self.gender = self.fill(self.gender,
                                                re.search("性别(.*?)(2018-05-08左侧|2018-04-16左侧)", content).group(1))
                    except Exception:
                        copyfile(self.file_path, "../../../output/errr/{}".format(self.file_path.split("\\")[-1]))
                        print("gender", self.file_path)
                        return 0
            # 提取 date
            try:
                self.date = self.fill(self.date, re.search("检查日期(.*?)打印日期", content).group(1))
            except Exception:
                try:
                    self.date = self.fill(self.date,
                                          re.search("(检查日期|时间)(.*?)(左侧|右侧|LCCA-IMT|2D及M型)", content).group(2))
                except Exception:
                    try:
                        self.date = self.fill(self.date, re.search("(性别男|性别女)(.*?)左侧CCA_IMT(mm)", content).group(1))
                    except Exception:
                        copyfile(self.file_path,
                                 "../../../output/errr/picture_no_string/{}".format(self.file_path.split("\\")[-1]))
                        with open("../../../output/err/err.txt", "a") as f:
                            f.write("{}{}\n".format(self.file_path, "图片非字符"))
                        print("date", self.file_path)
                        return 0
            if "LCCA-IMT" in content:
                try:
                    if content.index("LCCA-IMT") < content.index("RCCA-IMT"):
                        left = re.search("LCCA-IMT(.*?)RCCA-IMT", content).group(1)
                        right = re.search("RCCA-IMT(.*?)超声印象", content).group(1)
                    else:
                        right = re.search("RCCA-IMT(.*?)LCCA-IMT", content).group(1)
                        left = re.search("LCCA-IMT(.*?)超声印象", content).group(1)
                except Exception:
                    copyfile(self.file_path, "../../../output/errr/xindong/{}".format(self.file_path.split("\\")[-1]))
                    print(" no  RCCA-IMT", self.file_path)
            else:
                try:
                    if "右侧" in content and "左侧" in content:
                        if content.index("左侧") < content.index("右侧"):
                            left = re.search("左侧(.*?)右侧", content).group(1)
                            right = re.search("右侧(.*?)(超声印象|检查医生|报告医生)", content).group(1)
                        else:
                            left = re.search("右侧(.*?)左侧", content).group(1)
                            right = re.search("左侧(.*?)(超声印象|检查提示)", content).group(1)
                    else:
                        try:
                            left = re.search("右侧(.*?)右侧", content).group(1)
                            content = re.sub(left, "", content)
                            right = re.search("右侧(.*?)(超声印象|检查提示|报告医生)", content).group(1)
                        except Exception:
                            try:
                                left = re.search("左侧(.*?)左侧", content).group(1)
                                content = re.sub(left, "", content)
                                right = re.search("左侧(.*?)(超声印象|检查提示|报告医生)", content).group(1)
                            except Exception:
                                try:
                                    left = re.search("CCA-IMT（mm）(.*?)CCA-IMT（mm）", content).group(1)
                                    content = re.sub(left, "", content)
                                    right = re.search("CCA-IMT（mm）(.*?)(超声印象|检查医生|报告医生)", content).group(1)
                                except Exception:
                                    copyfile(self.file_path,
                                             "../../../output/errr/xindong/{}".format(self.file_path.split("\\")[-1]))
                                    print("left or right", self.file_path)
                                    return 0
                except Exception:
                    copyfile(self.file_path, "../../../output/errr/xindong/{}".format(self.file_path.split("\\")[-1]))
                    print("left or right", self.file_path)
                    return 0
                # print(left)
                # print(right)




            # 提取 CCA_IMT_right
            try:
                try:
                    self.CCA_IMT_right = (self.fill(self.CCA_IMT_right[0], re.search("近段(.*?)中段", right).group(1)),
                                          self.fill(self.CCA_IMT_right[1], re.search("中段(.*?)远段", right).group(1)),
                                          # self.fill(self.CCA_IMT_right[2], re.search("远段(.*?)增厚"), right).group(1))
                                          self.fill(self.CCA_IMT_right[2], re.search("远段(.*?)(数量|斑块)", right).group(1)))
                except Exception:
                    self.CCA_IMT_right = (self.fill(self.CCA_IMT_right[0], re.search("远段(.*?)中段", right).group(1)),
                                          self.fill(self.CCA_IMT_right[1], re.search("中段(.*?)近段", right).group(1)),
                                          self.fill(self.CCA_IMT_right[2], re.search("近段(.*?)(数量|斑块)", right).group(1)))
            except Exception:
                copyfile(self.file_path, "../../../output/errr/{}".format(self.file_path.split("\\")[-1]))
                print("CCA_IMT_right", self.file_path)
                return 0
            # 提取 plaques_count_right
            try:
                self.plaques_count_right = self.fill(self.plaques_count_right,
                                                     re.search("数量（1=单发，2=多发）[(.*?)]", right).group(1))
            except Exception:
                try:
                    self.plaques_count_right = self.fill(self.plaques_count_right,
                                                         re.search("数量1.无2.单发3.多发(.*?)(最大者长度|最长者长度)", right).group(1))
                except Exception:
                    try:
                        self.plaques_count_right = self.fill(self.plaques_count_right,
                                                             re.search(
                                                                 "(数量（1=单发，2=多发）|数量(1=单发，2=多发))(.*?)(最大者长度|最长者长度)",
                                                                 right).group(
                                                                 3))
                    except Exception:
                        copyfile(self.file_path, "../../../output/errr/geshi/{}".format(self.file_path.split("\\")[-1]))
                        print("plaques_count_right", self.file_path)
                        return 0
            # 提取 largest_plaque_width_right
            try:
                self.largest_plaque_width_right = self.fill(self.largest_plaque_width_right,
                                                            re.search("(最大者长度|最长者长度)(.*?)mm", right).group(2))
            except Exception:
                try:
                    self.largest_plaque_width_right = self.fill(self.largest_plaque_width_right,
                                                                re.search("(最大者长度|最长者长度)(.*?)(最大者厚度|最大厚度|最大着厚度)",
                                                                          right).group(2))
                except Exception:
                    try:
                        self.largest_plaque_width_right = self.fill(self.largest_plaque_width_right,
                                                                    re.search("最大者长度(.*?)3.3", right).group(1))
                    except Exception:
                        copyfile(self.file_path, "../../../output/errr/{}".format(self.file_path.split("\\")[-1]))
                        print("largest_plaque_width_right", self.file_path)
                        return 0
            # 提取 largest_plaque_depth_right
            try:
                self.largest_plaque_depth_right = self.fill(self.largest_plaque_depth_right,
                                                            re.search("最大者厚度(.*?)mm。", right).group(2))
            except Exception:
                try:
                    self.largest_plaque_depth_right = self.fill(self.largest_plaque_depth_right,
                                                                re.search("(最大厚度|最大着厚度|最大者厚度)(.*?)形态", right).group(2))
                except Exception:
                    try:
                        self.largest_plaque_depth_left = self.fill(self.largest_plaque_depth_left,
                                                                   re.search("最大者厚度(.*?)回声", right).group(
                                                                       1))
                    except Exception:
                        copyfile(self.file_path, "../../../output/errr/{}".format(self.file_path.split("\\")[-1]))
                        print("largest_plaque_depth_left", self.file_path)
                        return 0
            # 提取 plaque_shape_right
            try:
                self.plaque_shape_right = self.fill(self.plaque_shape_right,
                                                    re.search("形态（1=规则型，2=不规则型）[(.*?)]", right).group(1))
            except Exception:
                try:
                    self.plaque_shape_right = self.fill(self.plaque_shape_right,
                                                        re.search("形态1.规则型2.不规则型(.*?)有无溃疡斑块", right).group(1))
                except Exception:
                    try:
                        self.plaque_shape_right = self.fill(self.plaque_shape_right,
                                                            re.search(
                                                                "(形态（1=规则型，2=不规则型|形态（1=规则型，2=不规则型）|形态（1=规则型，2=不规则）|形态（1=规则型，2=不规则型。）|形态（1=规则，2=不规则）)(.*?)(是否溃疡|是否溃疡型)",
                                                                right).group(2).replace("）", ""))
                    except Exception:
                        copyfile(self.file_path, "../../../output/errr/{}".format(self.file_path.split("\\")[-1]))
                        print("plaque_shape_right", self.file_path)
                        return 0
            # 提取 plaque_is_ulcer_right
            try:
                self.plaque_is_ulcer_right = self.fill(self.plaque_is_ulcer_right,
                                                       re.search("是否溃疡型（0=否，1=是）[(.*?)]", right).group(1))
            except Exception:
                try:
                    self.plaque_is_ulcer_right = self.fill(self.plaque_is_ulcer_right,
                                                           re.search("有无溃疡斑块1.无2.有(.*?)狭窄程度", right).group(1))
                except Exception:
                    try:
                        self.plaque_is_ulcer_right = self.fill(self.plaque_is_ulcer_right,
                                                               re.search(
                                                                   "(是否溃疡型（0否，1=是）|是否溃疡型（0=否，1=是）|是否溃疡（0=否，1=是）|是否溃疡型A3（0=否，1=是）)(.*?)质地",
                                                                   right).group(2))
                    except Exception:
                        copyfile(self.file_path, "../../../output/errr/{}".format(self.file_path.split("\\")[-1]))
                        print("plaque_is_ulcer_right", self.file_path)
                        return 0
            # 提取 plaque_texture_right
            try:
                self.plaque_texture_right = self.fill(self.plaque_texture_right,
                                                      re.search("质地（A1=均质低回声，A2=均质等回声，A3均质强回声，B=不均质）[(.*?)]",
                                                                right).group(1))
            except Exception:
                try:
                    self.plaque_texture_right = self.fill(self.plaque_texture_right,
                                                          re.search("1.强回声2.中等回声3.低回声4.不均匀回声(.*?)形态1.规则型", right).group(
                                                              1))
                except Exception:
                    try:
                        self.plaque_texture_right = self.fill(self.plaque_texture_right,
                                                              re.search(
                                                                  "(质地（A1=均质低回声，A2=均质等回声，A3=均质强回声，B=不均匀）|质地（A1均质低回声，A2均质等回声，A3=均质强回声，B=不均质）|质地（A1=均质低回声，A2=均质等回声，A3=均质强回声，B=不均质）|质地（A1=均质低回声，A2=均质等回声，A3均质强回声，B=不均质）\[|质地（A1=均质低回声，A2=均质等1回声，A3=均质强回声，B=不均质）|质地（A1=均质低回声，A2=均质等回声，A3=不均质强回声，B=不均质）|质地（A1=均质低回7声，A2=均质等回声，A3=均质强回声，B=不均质）)(.*?)(管腔直径狭窄率|]。管腔直径狭窄率|狭窄部位)",
                                                                  right).group(
                                                                  2))
                    except Exception:
                        copyfile(self.file_path, "../../../output/errr/{}".format(self.file_path.split("\\")[-1]))
                        print("plaque_texture_right", self.file_path)
                        return 0
            # 提取 DS_right
            try:
                self.DS_right = self.fill(self.DS_right,
                                          re.search("(管腔直径狭窄率％|管腔直径狭窄率%)(.*?)狭窄部位", right).group(2))
            except Exception:
                try:
                    self.DS_right = self.fill(self.DS_right, re.search("狭窄程度或闭塞部位(.*?)(检查结果|左侧)", right).group(1))
                except Exception:
                    try:
                        self.DS_right = self.fill(self.DS_right, re.search("管腔直径狭窄率(.*?)(%狭窄部位|狭窄部位)", right).group(1))
                    except Exception:
                        try:
                            self.DS_right = self.fill(self.DS_right, re.search("(狭窄程度或闭塞部位|狭窄部位)(.*?)", right).group(1))
                        except Exception:
                            copyfile(self.file_path, "../../../output/errr/{}".format(self.file_path.split("\\")[-1]))
                            print("DS_right", self.file_path)
                            return 0
            # print(right)
            # 提取 location_right
            try:
                self.location_right = self.fill(self.location_right, re.search("狭窄部位(.*?)超声印象", right).group(1))
            except Exception:
                try:
                    self.location_right = self.fill(self.location_right, re.search("狭窄部位(.*)", right).group(1))
                except Exception:
                    try:
                        self.location_right = self.fill(self.location_right,
                                                        re.search("狭窄程度或闭塞部位(.*?)检查结果", right).group(1))
                    except Exception:
                        try:
                            self.location_right = self.fill(self.location_right,
                                                            re.search("狭窄程度或闭塞部位(.*?)左侧", content).group(1))
                        except Exception:
                            copyfile(self.file_path, "../../../output/errr/{}".format(self.file_path.split("\\")[-1]))
                            print("location_right", self.file_path)
                            return 0
            # 提取 comments
            # print(content)
            try:
                self.comments = self.fill(self.comments, re.search("超声印象;(.*?)心血管病高危人群早期筛查与综合干预项目", content).group(1))
            except Exception:
                try:
                    self.comments = self.fill(self.comments, re.search("超声印象(.*?)(报告医生|报告医师)", content).group(1))
                except Exception:
                    try:
                        self.comments = self.fill(self.comments, re.search("超声印像(.*?)报告医生", content).group(1))
                    # 添加  A3管腔直径狭窄率%狭窄部位(.*?)报告医生白妍，right
                    except Exception:
                        try:
                            self.comments = self.fill(self.comments, re.search("检查提示(.*?)检查医生", content).group(1))
                        except Exception:
                            try:
                                self.comments = self.fill(self.comments, re.search("狭窄部位(.*)", right).group(1))
                            except Exception:
                                try:
                                    self.comments = self.fill(self.comments, re.search("狭窄部位(.*?)报告医生", content).group(1))
                                    # print("1", self.comments)
                                except Exception:
                                    copyfile(self.file_path,
                                             "../../../output/errr/{}".format(self.file_path.split("\\")[-1]))
                                    print("comments", self.file_path)
                                    return 0






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
            f.Close()
        with open("../../../output/output.txt", "a")as output:
            output.write(
                "{} {} {} {} {} {} {} {} {} {} {} {} {} {} {} \n".format(self.name, self.ID, self.gender, self.date,
                                                                         self.CCA_IMT_right, self.plaques_count_right,
                                                                         self.largest_plaque_width_right,
                                                                         self.largest_plaque_depth_right,
                                                                         self.plaque_shape_right,
                                                                         self.plaque_is_ulcer_right,
                                                                         self.plaque_texture_right, self.DS_right,
                                                                         self.location_right, self.comments,
                                                                         self.doctor))
        # print(content)
        # print(self.location_right)
        # print(self.doctor)
        # print(content)


if __name__ == "__main__":
    word = Dispatch("Word.Application")
    test = UltraSound(
        # " G:\\机器学习\\doc_conversion\\data\\20190722\\IV期\\Long_fu_survey_Carotid_ultrasound\\2202\\G220203443.doc", word)
        # "G:\\机器学习\\doc_conversion\\data\\20190722\\IV期\\Long_fu_survey_Carotid_ultrasound\\2202\\G220203188.doc", word)
        r"F:\data\20190722\IV期\High_risk_Carotid_ultrasound\2101\G2101400022.doc",
        word)
