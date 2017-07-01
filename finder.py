# -*- coding: cp949 -*-
import postfile
import simplejson
import urllib
import urllib2
import sys
import xlwt
import os
import time

ss = sys.argv[1]
arr = []
spray = []
j = 0
print ss+u"를 탐색합니다."


host = "www.virustotal.com"
url = "https://www.virustotal.com/vtapi/v2/file/report"
selector = "https://www.virustotal.com/vtapi/v2/file/scan" #검색대상
fields = [("apikey", "5ad70c8065f80b022e92e73f6643778b94b80edc9bcfab019d1a3dcd83590177")] #API키 삽입
#선언부

path_dir = ss
file_list = os.listdir(path_dir)
arr = file_list
#파일 경로 설정
workbook = xlwt.Workbook(encoding='utf-8')

for i in range (len(arr)):
    filename = arr[i]
    print filename + u"파일을 검사합니다"

    File_to_send = open(arr[i],'rb').read()

    files = [("file", arr[i], File_to_send)]
    file_send = postfile.post_multipart(host, selector, fields, files)
    dict_data = simplejson.loads(file_send)
    resource = dict_data.get("resource", {})
    parameters = {"resource": resource, "apikey": "5ad70c8065f80b022e92e73f6643778b94b80edc9bcfab019d1a3dcd83590177"}

    data = urllib.urlencode(parameters)
    req = urllib2.Request(url, data)

    response = urllib2.urlopen(req)
    resource_data = response.read()

    result = simplejson.loads(resource_data)
    spray = result['scans']
    #데이터 처리

    ############################### 엑셀 처리 부분 ####################################

    workbook.default_style.font.heignt = 20 * 11
    #시트 생성 & 기본 폰트 설정

    xlwt.add_palette_colour("lightgray", 0x21)
    workbook.set_colour_RGB(0x21, 216, 216, 216)
    xlwt.add_palette_colour("lightgreen", 0x22)
    workbook.set_colour_RGB(0x22, 216, 228, 188)
    file = arr[i]

    worksheet = workbook.add_sheet(file + "", cell_overwrite_ok=True)
    col_width_0 = 256 * 21
    col_width_1 = 256 * 13
    col_width_2 = 256 * 21
    col_width_3 = 256 * 13
    col_width_4 = 256 * 13

    col_height_content = 48

    worksheet.col(0).width = col_width_0
    worksheet.col(1).width = col_width_1
    worksheet.col(2).width = col_width_2
    worksheet.col(3).width = col_width_3
    worksheet.col(4).width = col_width_4


    # 항목 스타일 설정
    list_style = "font:height 180,bold on; pattern: pattern solid, fore_color lightgray; align: wrap on, vert centre, horiz center"

    # 엑셀에 항목 입력


    worksheet.write_merge(0, 0, 0, 4, ss+"/"+arr[i], xlwt.easyxf(list_style)) #셀 병합
    worksheet.write(1,0,'sha-1',xlwt.easyxf(list_style))
    worksheet.write_merge(1, 1, 1, 4, result['sha1'])

    worksheet.write(2, 0, u"백신", xlwt.easyxf(list_style))
    worksheet.write(2, 1, u"탐지", xlwt.easyxf(list_style))
    worksheet.write(2, 2, u"백신버전", xlwt.easyxf(list_style))
    worksheet.write(2, 3, u"결과", xlwt.easyxf(list_style))
    worksheet.write(2, 4, u"패턴날짜", xlwt.easyxf(list_style))


    keys = []
    head = []
    keys = spray.keys()
    head = spray.keys()
    res = []
    res = list(spray.values())


    for i in keys:
        worksheet.write(j+3,0, head[j])  # 백신이름

        vaccinedetected = (spray[i]['detected'])
        worksheet.write(j+3,1, vaccinedetected) #탐지

        vaccineversion = (spray[i]['version'])
        worksheet.write(j+3,2, vaccineversion) #백신이름

        vaccineversion = (spray[i]['result'])
        if(vaccineversion == None):
            worksheet.write(j + 3, 3, "NULL!")  # 결과
        else:
            worksheet.write(j + 3, 3, vaccineversion)

        vaccineversion = (spray[i]['update'])
        worksheet.write(j + 3, 4, vaccineversion)  # 업데이트

        j = j + 1

        workbook.save("result.xls")



    time.sleep(3)
    print("complete")
    j = 0 # 새로운파일 카운트 초기화
    spray = []