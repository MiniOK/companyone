import os

import xlrd

DIR = "../../../data/20190722/"

# Get all files in list
word_documents = []
excel_spreadsheets = []
pdfs = []

for version in os.listdir(DIR):
    sub_dir = "{}/{}/".format(DIR, version)
    for category in os.listdir(sub_dir):
        ssub_dir = sub_dir + category + "/"
        for item in os.listdir(ssub_dir):
            sssub_dir = ssub_dir + item + "/"
            for file in os.listdir(sssub_dir):
                if file.split(".")[-1].lower() in ["doc", "docx"]:
                    word_documents.append(sssub_dir + file)
                elif file.split(".")[-1].lower() in ["xls", "xlsx"]:
                    excel_spreadsheets.append(sssub_dir + file)
                elif file.split(".")[-1].lower() in ["pdf"]:
                    pdfs.append(sssub_dir + file)


def _read_xl(fn):
    # open the work book
    wb = xlrd.open_workbook(fn)

    # open the sheet you want to read its cells
    sht = wb.sheet_by_index(0)
    # if sht.cell(0,0).value == "颈动脉超声检查报告":
    for r in range(sht.nrows):
        for c in range(sht.ncols):
            cell_col1 = sht.cell(rowx=r, colx=c).value
            # if cell_col1 != "":
            #     if r in [10, 19]:
            #         print(fn, r, c, cell_col1)
            print(r,c,cell_col1)
        return sht.cell(0,1)
    return {}


# for file in excel_spreadsheets[:200]:
    # _read_xl(file)
a  = _read_xl('M:\\PycharmProjects\\doc_conversion\\data\\20190722\\IV期\\High_risk_Carotid_ultrasound\\5106\\G5106400033.xlsx')