import json
import xlwt
import time

file = open("D:\logs\search.json", encoding='utf-8')  # 设置以utf-8解码模式读取文件，encoding参数必须设置，
setting = json.load(file)
data = setting["data"]
needData = data["data"]
flightList = []
for flight in needData:
    flightList.append(flight)
# 创建excel对象
workbook = xlwt.Workbook()
worksheet = workbook.add_sheet('sheet1')  # 添加一个表
pattern = xlwt.Pattern()  # Create the Pattern
pattern.pattern = xlwt.Pattern.SOLID_PATTERN  # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
pattern.pattern_fore_colour = 5
style = xlwt.XFStyle()  # 初始化样式
style.pattern = pattern  # Add Pattern to Style
i = 0
for flight in flightList:
    #将数据写入文件,i是enumerate()函数返回的序号数
    worksheet.write(i, 0, flight["dptCity"], style)
    worksheet.write(i, 1, flight["dptCity"], style)
    worksheet.write(i, 2, flight["arrCity"], style)
    worksheet.write(i, 3, flight["dptAirport"], style)
    worksheet.write(i, 4, flight["arrAirport"], style)
    worksheet.write(i, 5, flight["carrier"], style)
    worksheet.write(i, 6, flight["cabin"], style)
    worksheet.write(i, 7, flight["price"], style)
    worksheet.write(i, 8, flight["sharecodeLimit"], style)
    worksheet.write(i, 9, flight["sharecodeForbidden"], style)
    worksheet.write(i, 10, flight["minPreDays"], style)
    worksheet.write(i, 11, flight["maxPreDays"], style)
    worksheet.write(i, 12, flight["arrAirport"], style)
    worksheet.write(i, 13, flight["ticketStart"], style)
    worksheet.write(i, 14, flight["ticketEnd"], style)
    worksheet.write(i, 15, flight["dptAirport"], style)
    worksheet.write(i, 16, flight["flightNumLimit"], style)
    worksheet.write(i, 17, flight["flightNumForbidden"], style)
    worksheet.write(i, 18, flight["fareBasis"], style)
    worksheet.write(i, 19, flight["inTravelTime"], style)
    worksheet.write(i, 20, flight["forbiddenTravelTime"], style)
    worksheet.write(i, 21, flight["forbiddenTravelDate"], style)
    worksheet.write(i, 22, flight["owSeasonValidDate"], style)
    worksheet.write(i, 23, flight["weekLimit"], style)
    worksheet.write(i, 24, flight["ifIndividual"], style)
    worksheet.write(i, 25, flight["ifGroup"], style)
    worksheet.write(i, 26, flight["ifRoundTrip"], style)
    worksheet.write(i, 27, flight["ifOneWay"], style)
    worksheet.write(i, 28, flight["officeId"], style)
    worksheet.write(i, 29, flight["reFundRules"], style)
    worksheet.write(i, 30, flight["changeRules"], style)
    worksheet.write(i, 31, flight["signRules"], style)
    worksheet.write(i, 32, flight["dataSource"], style)
    worksheet.write(i, 33, time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(flight["createTime"])), style)
    worksheet.write(i, 34, time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(flight["updateTime"])), style)
    worksheet.write(i, 35, flight["trace"], style)
    i += 1
workbook.save(r'D:\logs\flightExcel.xls')