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
print ss+u"�� Ž���մϴ�."


host = "www.virustotal.com"
url = "https://www.virustotal.com/vtapi/v2/file/report"
selector = "https://www.virustotal.com/vtapi/v2/file/scan" #�˻����
fields = [("apikey", "5ad70c8065f80b022e92e73f6643778b94b80edc9bcfab019d1a3dcd83590177")] #APIŰ ����
#�����

path_dir = ss
file_list = os.listdir(path_dir)
arr = file_list
#���� ��� ����
workbook = xlwt.Workbook(encoding='utf-8')

for i in range (len(arr)):
    filename = arr[i]
    print filename + u"������ �˻��մϴ�"

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
    #������ ó��

    ############################### ���� ó�� �κ� ####################################

    workbook.default_style.font.heignt = 20 * 11
    #��Ʈ ���� & �⺻ ��Ʈ ����

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


    # �׸� ��Ÿ�� ����
    list_style = "font:height 180,bold on; pattern: pattern solid, fore_color lightgray; align: wrap on, vert centre, horiz center"

    # ������ �׸� �Է�


    worksheet.write_merge(0, 0, 0, 4, ss+"/"+arr[i], xlwt.easyxf(list_style)) #�� ����
    worksheet.write(1,0,'sha-1',xlwt.easyxf(list_style))
    worksheet.write_merge(1, 1, 1, 4, result['sha1'])

    worksheet.write(2, 0, u"���", xlwt.easyxf(list_style))
    worksheet.write(2, 1, u"Ž��", xlwt.easyxf(list_style))
    worksheet.write(2, 2, u"��Ź���", xlwt.easyxf(list_style))
    worksheet.write(2, 3, u"���", xlwt.easyxf(list_style))
    worksheet.write(2, 4, u"���ϳ�¥", xlwt.easyxf(list_style))


    keys = []
    head = []
    keys = spray.keys()
    head = spray.keys()
    res = []
    res = list(spray.values())


    for i in keys:
        worksheet.write(j+3,0, head[j])  # ����̸�

        vaccinedetected = (spray[i]['detected'])
        worksheet.write(j+3,1, vaccinedetected) #Ž��

        vaccineversion = (spray[i]['version'])
        worksheet.write(j+3,2, vaccineversion) #����̸�

        vaccineversion = (spray[i]['result'])
        if(vaccineversion == None):
            worksheet.write(j + 3, 3, "NULL!")  # ���
        else:
            worksheet.write(j + 3, 3, vaccineversion)

        vaccineversion = (spray[i]['update'])
        worksheet.write(j + 3, 4, vaccineversion)  # ������Ʈ

        j = j + 1

        workbook.save("result.xls")



    time.sleep(3)
    print("complete")
    j = 0 # ���ο����� ī��Ʈ �ʱ�ȭ
    spray = []