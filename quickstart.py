## xlwt quickstart

import xlwt
import math
from datetime import datetime

##style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on', num_format_str='#,##0.00')
##style1 = xlwt.easyxf(num_format_str='D-MMM-YY')
##
##wb = xlwt.Workbook()
##ws = wb.add_sheet('TestSheet1')
##
##ws.write(0, 0, 1234.56, style0)
##ws.write(1, 0, datetime.now(), style1)
##ws.write(2,0,1)
##ws.write(2,1,1)
##ws.write(2,2,xlwt.Formula("A3+B3"))
##
##wb.save('example.xls')
##

wb = xlwt.Workbook()

## generate cam outline
ws = wb.add_sheet('Cam outline')

polarx = 0
polarmax = 2 * math.pi
pitchradius = 18
amplitude = 1.40
writerow = 0

while polarx < polarmax:
    r = pitchradius + amplitude * math.sin(polarx * 8 - math.pi / 2)
    x = r * math.cos(polarx)
    y = r * math.sin(polarx)
    ws.write(writerow,0,x)
    ws.write(writerow,1,y)
    writerow += 1
    polarx += 0.1

wb.save('CamOutline.xls')
    
