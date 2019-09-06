# Conf

# Data

# output
# scripts
* conversion
    * main是doc入口
    * main2是xls入口
    * 以上两个均仅供测试用
* load_file
    * doc 
        * 将word文档转化为字符串，去掉特殊字符、制表符
        * 用正则表达式扣取要用的字段，由于文档大小不一、形态各异，写了许多规则。尚不全面，且会有很多细枝末节的问题
        * 待优化
    * xls
        * 根据表格样式从单元格提取信息。
        * 待处理：开头有医院或机构名字的
# temp
