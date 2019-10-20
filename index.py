from openpyxl import *
import datetime

wb = load_workbook(r".\HasilPemetaanIndustri.xlsx")


variable = {}
merek = []
type  = []
kapasitas = []
for sheet in wb.worksheets:
    checkMerek = False
    for i in range(15,25):
        if sheet['B'+str(i)].value == 'Merek':
            checkMerek = True
            continue
        
        if sheet['A'+str(i)].value == 'V.JUMLAH TENAGA KERJA' or sheet['A'+str(i)].value == 'V.JUMLAH TENAGA KERJA (NIHIL)' or sheet['A'+str(i)].value == 'V.JUMLAH TENAGA KERJA(NIHIL)':
            break

        if checkMerek and sheet['B'+str(i)].value != None and not sheet['B'+str(i)].value in  merek:
            merek.append(sheet['B'+str(i)].value)
        
        if checkMerek and sheet['C'+str(i)].value != None and not sheet['C'+str(i)].value in  type:
            type.append(sheet['C'+str(i)].value)
        
        if checkMerek and sheet['D'+str(i)].value != None and not sheet['C'+str(i)].value in  kapasitas:
            kapasitas.append(sheet['D'+str(i)].value)

print('ini merek: ')
print(merek)
print('\n\nini type: ')
print(type)
print('\n\nini kapasitas: ')
print(kapasitas)