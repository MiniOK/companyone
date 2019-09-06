import datetime

startDate = "2018.10.01"
endDate = "2018-10-31"

new = startDate.replace(".","-")
print(new)

###字符转化为日期
startTime = datetime.datetime.strptime(startDate, '%Y-%m-%d').time()
endTime = datetime.datetime.strptime(endDate, '%Y-%m-%d').time()

