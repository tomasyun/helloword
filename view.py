# from django.http import HttpResponse,JsonResponse
from django.shortcuts import render
from hellosql.models import Users
import json
import xlwt

# def hello(request):
# return HttpResponse("hello world !");

def hello(request):
    context = {}
    # context['hello'] = 'Hello World!'
    all_entries = Users.objects.all()
    return render(request, 'hello.html', {'all_entries': all_entries})
    # return render(request, 'hello.html', context)


def index(request):
    context = {}
    context['index'] = 'welcome to here!'
    return render(request, 'index.html', context)


def analysis(request):
    file = open("D:\logs\search.json", encoding='utf-8')  # 设置以utf-8解码模式读取文件，encoding参数必须设置，
    setting = json.load(file)
    data = setting["data"]
    needData = data["data"]
    flightList = []
    for flight in needData:
        flightList.append(flight)
    return render(request, 'detail.html', {'flightList': flightList})

# 写入excel文件
def export(data):
    # 创建excel对象
    workbook = xlwt.Workbook()
    worksheet = workbook.add_sheet('sheet1')  # 添加一个表
    pattern = xlwt.Pattern()  # Create the Pattern
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN  # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
    pattern.pattern_fore_colour = 5
    style = xlwt.XFStyle()  # Create the Pattern
    style.pattern = pattern  # Add Pattern to Style
    for i, p in enumerate(data):
        # 将数据写入文件,i是enumerate()函数返回的序号数
        for j, q in enumerate(p):
            worksheet.write(i, j, q, style)
    workbook.save('D:\logs\flightExcel.xls')