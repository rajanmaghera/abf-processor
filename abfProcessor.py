import pyabf
import xlsxwriter as xs
import matplotlib.pyplot as plt
import datetime
import os
from xlsxwriter.utility import xl_rowcol_to_cell as xl

folderName = "Input"
dir = os.listdir(folderName)

ver = 0
workbook = xs.Workbook("GvProcessed3000.xlsx")
sheet = workbook.add_worksheet("export")

sheet.set_row(4, 50)

titleCell = workbook.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'text_wrap': True,
        'font_size': 9
    })


for filename in dir:

    try:
        abf = pyabf.ABF(folderName + "/" + filename)
    except:
        print(f"File {filename} Failed, going to next one" )
        continue

    array = []
    delta = []
    peaks = []
    peak = 1

    inPeak = False
    max = 0

    delta.append(0)

    for x in range(1, len(abf.sweepY)):
        delta.append(abs(abf.sweepY[x] - abf.sweepY[x-1]))
        
    for x in range (len(abf.sweepY)):
        if inPeak == False and delta[x] > 6:
            inPeak = True
        
        if inPeak == True and delta[x] > 6:
            if delta[x] > delta[max]:
                max = x
        
        if inPeak == True and delta[x] < 6:
            peaks.append(x)
            inPeak = False

    if len(peaks) < 3:
        peak = 0


    spot = int(peaks[peak]) - 3

    print(spot)


    for i in range(abf.sweepCount):
        abf.setSweep(abf.sweepCount - i - 1)
        array.append([abf.sweepC[spot-int(0.0040/0.0002)],abf.sweepY[spot+int(0.006/0.0002)]])

    print(array)

    col = 0
    rowAdd = (int((array[0][0] + 210)/10))



    sheet.merge_range(2, 2+(ver*6), 3, 6+(ver*6), filename, titleCell)
    sheet.write_string(4, 2+(ver*6), '''Voltage
(mV)''', titleCell)
    sheet.write_string(4, 3+(ver*6), '''Tail Current Magnitude
(pA)''', titleCell)
    sheet.write_string(4, 4+(ver*6), '''Normalized Tail Current''', titleCell)    
    sheet.write_string(4, 5+(ver*6), '''Boltzmann Model Fit''', titleCell)    
    sheet.write_string(4, 6+(ver*6), '''Squared Difference''', titleCell)    
    sheet.write_string(33, 2+(ver*6), "MAX", titleCell)
    sheet.write_string(34, 2+(ver*6), "MIN", titleCell)
    sheet.write_string(35, 5+(ver*6), "z->", titleCell)
    sheet.write_string(36, 5+(ver*6), "V1/2->", titleCell)
    sheet.write_string(37, 5+(ver*6), "Sum of Squares", titleCell)
    sheet.write_number(35, 6+(ver*6), 2.9)
    sheet.write_number(36, 6+(ver*6), -35)

    for row, data in enumerate(array):
        sheet.write_row(row+rowAdd, col+2+(6*ver), data)
        sheet.write_formula(row+rowAdd, col+4+(6*ver), f"={xl(row+rowAdd, col+3+(6*ver))}/{xl(33, 3+(ver*6))}")
        sheet.write_formula(row+rowAdd, col+5+(6*ver), f"=1/(1+EXP(-({xl(35, 6+(ver*6))}*96845*({xl(row+rowAdd, col+2+(6*ver))}-{xl(36, 6+(ver*6))})/(8.3145*298*1000))))")
        sheet.write_formula(row+rowAdd, col+6+(6*ver), f"=({xl(row+rowAdd, col+5+(6*ver))}-{xl(row+rowAdd, col+4+(6*ver))})^2")

    sheet.write_formula(33, 3+(ver*6), f"=MAX({xl(5, 3+(ver*6))}:{xl(31, 3+(ver*6))})")
    sheet.write_formula(34, 3+(ver*6), f"=MIN({xl(5, 3+(ver*6))}:{xl(31, 3+(ver*6))})")
    sheet.write_formula(37, 6+(ver*6), f"=SUM({xl(5, 6+(ver*6))}:{xl(31, 6+(ver*6))})")
    sheet.write_formula(41, 6+(ver*6), f"=(8.3145*298*1000)/(96845*{xl(35, 6+(ver*6))})")
   
    ver += 1

workbook.close()
