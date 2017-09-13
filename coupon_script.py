import xlwt
import random
from datetime import datetime
from datetime import timedelta, date

style0 = xlwt.easyxf('font: name Times New Roman, bold on',
    num_format_str='#,##0.00')
style1 = xlwt.easyxf(num_format_str='D-MMM-YY')

wb = xlwt.Workbook()
ws = wb.add_sheet('Sheet1')

ws.write(0, 0, 'Code', style1)
ws.write(0, 1, 'Type', style1)
ws.write(0, 2, 'Expiring_Date_Time', style1)
#ws.write(0, 3, 'Charge_Point_Ref', style1)
Type = ['Bcode','Ucode']
def daterange(start_date, end_date):
    for n in range(int ((end_date - start_date).days)):
        yield start_date + timedelta(n)

start_date = date(2016, 1, 1)
end_date = date(2017, 12, 30)
dates = []
for single_date in daterange(start_date, end_date):
	dates.append(single_date.strftime("%Y/%m/%d %H:%M:%S"))

for x in range(1,50):

	ws.write(x, 0,random.randint(10043073,23343273))
	ws.write(x, 1,random.choice(Type))
	ws.write(x, 2,random.choice(dates))
	#ws.write(x, 3,random.randint(3453,9342))

wb.save('coupons1.xls')