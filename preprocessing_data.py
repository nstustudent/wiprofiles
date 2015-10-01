# -*- coding: utf-8 -*-
import xlrd, datetime, csv

def xldate_to_datetime(xldate):
  tempDate = datetime.datetime(1900, 1, 1)
  deltaDays = datetime.timedelta(days=int(xldate))
  TheTime = (tempDate + deltaDays )
  return TheTime.strftime("%Y-%m-%d")

#calculate infections
name_sheet = 'Infection_Ekater_p2'
data = []
rd = xlrd.open_workbook('./data/inf_data.xls')
sheet = rd.sheet_by_name(name_sheet)
for rownum in range(sheet.nrows):
    if rownum > 0:
        row = sheet.row_values(rownum)
        date = datetime.datetime(*xlrd.xldate_as_tuple(row[3], rd.datemode))
        print (date.date())
        data.append(str(date.date()))

dict_data = {}
for elem in data:
    if elem in dict_data.keys():
        dict_data[elem] += 1
    else:
        dict_data[elem] = 1

writer = csv.writer(open('./data/'+name_sheet+'.csv','wb'))
for el in sorted(dict_data):
    temp = [el, dict_data[el]]
    writer.writerow(temp)





