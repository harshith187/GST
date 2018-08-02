import xlrd
from collections import OrderedDict
import simplejson as json
from datetime import datetime

wb = xlrd.open_workbook('/Users/harshith/projects/gst/converted_gstin.xlsx')
sh = wb.sheet_by_name('b2b')
print type(sh)


sheet_details = OrderedDict()

sheet_details['fp']='062018'
sheet_details['gstin']='29AKCPG8933G1Z0'
sheet_details['hash']='hash'
sheet_details['version']='GST2.2.6'

gstins =[]
for row in range(4,sh.nrows):
    row_values = sh.row_values(row)
    gstins.append(row_values[0])


data_list = []
for rownum in range(4, sh.nrows):
    data = OrderedDict()
    row_values = sh.row_values(rownum)
    data['ctin'] = row_values[0]

    invoice_list=[]
    invoice = OrderedDict()
    row_values = sh.row_values(rownum)
    invoice['inum'] = row_values[2]
    invoice_date = datetime.strptime(row_values[3],"%d-%b-%Y")
    invoice['idt']=datetime.strftime(invoice_date, '%d-%m-%Y')
    invoice['val']=row_values[4]
    invoice['pos']=row_values[5][:2]
    invoice['rchrg']=row_values[6]
    invoice['inv_typ']=row_values[8][:1]
    data3_list=[]
    
    data3=OrderedDict()
    row_values=sh.row_values(rownum)
    data3['num']=1201
    items_details=OrderedDict()
    row_values=sh.row_values(rownum)
    items_details['txval']=row_values[11]
    items_details['rt']=int(row_values[10][:2])
    items_details['camt']=''
    items_details['samt']=''
    items_details['csamt']=0

    data3['itm_det']=items_details
    invoice['itms']=data3_list
    data['inv'] = invoice_list
    sheet_details['b2b']=data_list
    
    data_list.append(data)
    invoice_list.append(invoice)
    data3_list.append(data3)
    

j = json.dumps(sheet_details)

with open('final.json', 'w') as f:
    f.write(j)