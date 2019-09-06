


with open('../output/finish-001.txt', 'r+', encoding='utf-8') as finish_txt:
    data=finish.read()
    print(data)
    finish.write("{}\t\t{}\t\t\t{}\t\t\t{}\t\t{}\t\t\t{}\t\t\t{}\t\t\t{}\t\t\t{}\t\t\t{}\t\t\t{}\t\t\t{}\t\t\t{}\t\t\t{}\t\t\t\n"
                 .format("name","ID","sex","date","v5","v6","v7","v8","v9","v10","v11","v12","v13","v14"))
    for i in range(10):
        finish.write("{}\t\t{}\t\t\t{}\t\t\t{}\t\t\t{}\t\t\t{}\t\t\t{}\t\t\t{}\t\t\t{}\t\t\t{}\t\t\t{}\t\t\t{}\t\t\t{}\t\t\t{}\t\t\t\n"
                     .format(1,2,3,4,5,6,7,8,9,10,11,12,13,14))