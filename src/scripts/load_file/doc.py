from win32com.client import Dispatch


def read_doc(file_dirs):
    word = Dispatch("Word.Application.8")
    word.Visible = 0
    output = {}
    for file_dir in file_dirs:
        f = word.Documents.Open(file_dir)
        # lines = []
        if len(f.Tables) > 1:
            print("opps")
        for t in f.Tables:
            for i in range(t.Rows.Count):
                for j in range(t.Columns.Count):
                    try:
                        print(i + 1, j + 1, t.Cell(i + 1, j + 1))
                    except Exception:
                        print(i+1, j+1, None)
            # table_text = t.__str__()
            # lines = table_text.split("\r\x07")
        # output[file_dir] = lines
        f.Close()
    return output

fn = 'M:\\PycharmProjects\\doc_conversion\\data\\20190722\\IV期\\Long_fu_survey_Carotid_ultrasound\\2202\\G220200878何艳华.doc'
read_doc([fn])
word = Dispatch("Word.Application.8")
f = word.Documents.Open(fn)