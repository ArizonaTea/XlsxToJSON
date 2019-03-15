import glob
from win32com.client import Dispatch
import xlrd
import os
from collections import OrderedDict
import simplejson as json


#Transfer xlsx TO xls

target_dir_xls=r"C:\Users\tliu\Desktop\Lipsticks_3_13\XLS"
if not os.path.exists(target_dir_xls):
    print('Target file path not exist, Creating...')
    os.mkdir(target_dir_xls)

for filename in glob.glob(r'C:\Users\tliu\Desktop\Lipsticks_3_13\*.xlsx'):
    xl = Dispatch('Excel.Application')
    wb = xl.Workbooks.Add(filename)
    wb.SaveAs(filename[:-1], FileFormat=56)
    xl.Quit()


#Trasfer xls to JSON

target_dir_json=r"C:\Users\tliu\Desktop\Lipsticks_3_13\JSON"
if not os.path.exists(target_dir_json):
    print('Target file path not exist, Creating...')
    os.mkdir(target_dir_json)


for filename in glob.glob(r'C:\Users\tliu\Desktop\Lipsticks_3_13\XLS\*.xls'):
    xlsxfile = xlrd.open_workbook(filename)
    sh = xlsxfile.sheet_by_index(0)
    jsonfile = filename + '.json'
    xls_list = []
    for rownum in range(1, sh.nrows):
        lips = OrderedDict()
        row_values = sh.row_values(rownum)
        lips['Name'] = row_values[0]
        lips['Description'] = row_values[1]
        lips['Ref Number'] = row_values[2]
        lips['Colour Code'] = row_values[3]
        lips['Colour Image'] = row_values[4]
        lips['Price'] = row_values[5]
        lips['Unit'] = row_values[6]
        lips['Product Image'] = row_values[7]
        lips['Purchase Link'] = row_values[8]
        lips['Key Words'] = row_values[9]
        xls_list.append(lips)

    j = json.dumps(xls_list)

    with open(jsonfile, 'w') as f:
        f.write(j)



