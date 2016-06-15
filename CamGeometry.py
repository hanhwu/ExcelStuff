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

cam_angle = 0
pitch_radius = 18
amplitude = 1.40
write_row = 0

while cam_angle < 2 * math.pi:
    r = pitch_radius + amplitude * math.sin(cam_angle * 8 - math.pi / 2)
    x = r * math.cos(cam_angle)
    y = r * math.sin(cam_angle)
    ws.write(write_row,0,x)
    ws.write(write_row,1,y)
    write_row += 1
    cam_angle += 0.01


roller_radius = 3
roller_min = pitch_radius - amplitude - roller_radius - 0.01
roller_max = pitch_radius + amplitude - roller_radius + 0.01

# test with arbitrary roller angle

roller_angle = math.pi / 6
testmin = roller_min
testmax = roller_max

guess = (testmin + testmax) / 2 # guess of roller center distance from cam center

while testmax - testmin > 0.001:
    contact_sweep_min = roller_angle - math.atan( roller_radius / guess ) + math.pi / 20
    contact_sweep_max = roller_angle + math.atan( roller_radius / guess ) - math.pi / 20
    sweep_increment = (contact_sweep_max - contact_sweep_min) / 1000
    contact_sweep = contact_sweep_min
    intersect = False
    while contact_sweep < contact_sweep_max:
        roller_edge = guess * math.cos(contact_sweep - roller_angle) + math.sqrt( roller_radius**2 - (guess * math.sin( contact_sweep - roller_angle ))**2 )
        cam_edge = guess + amplitude * math.sin(8 * contact_sweep - math.pi / 2)
        if roller_edge > cam_edge:
            intersect = True
            pressure_angle = math.pi - math.acos((roller_edge**2 - guess**2 - roller_radius**2) / ( -2 * guess * roller_radius))
            break

        contact_sweet += sweep_increment


    if intersect:
        testmax = guess
        guess = (testmin + testmax)/2
    else:
        testmin = guess
        guess = (testmin + testmax)/2
        
    print 'Current guess =', guess

print 'Roller angle (rad) =', roller_angle,', roller center R =', guess, 'pressure angle (rad) =', pressure_angle











    









wb.save('CamOutline.xls')
    
