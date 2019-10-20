# import xlsxwriter module
import xlsxwriter
import csv
import sys

data_set = 'water'

workbook = xlsxwriter.Workbook(data_set+'.xlsx')
# By default worksheet names in the spreadsheet will be
# Sheet1, Sheet2 etc., but we can also specify a name.
worksheet = workbook.add_worksheet(data_set)

Yt = []
data_sum = 0
with open(data_set+'.csv','rt')as f:
    data = csv.reader(f)
    for row in data:
        d = float(row[0])
        Yt.append(float(row[0]))
        data_sum = data_sum+d

print Yt
Ybar = data_sum/len(Yt)
#sys.exit(0)

total_row = len(Yt)
total_column = total_row
#print 'Total Row', total_row
#print 'Total Column', total_column
end_column = 0
r_d = 0
ac_list = []

for r in range(-1, total_row):
    if r == -1:
        worksheet.write(r+1, 0, 'Yt')
    else:
        worksheet.write(r+1, 0, Yt[r])


for r in range(-1, total_row):
    if r == -1:
        worksheet.write(r+1, 1, 'Yt-Ybar')
    else:
        worksheet.write(r+1, 1, Yt[r]-Ybar)

for r in range(-1, total_row):
    if r == -1:
        worksheet.write(r+1, 2, '(Yt-Ybar)^2')
    else:
        x = (Yt[r]-Ybar)**2
        worksheet.write(r+1, 2, x)
        r_d = x + r_d

lag = 1
for c in range(3, (2*total_column)+3):
    if c % 2 != 0:
        Yadd = ['-'] * lag + Yt
        Ylag = Yadd[0:len(Yadd)-lag]
        lag = lag + 1
    r_n = 0
    for r in range(-1, total_row):
        if r == -1:
            if c %2 != 0:
                name = 'Yt-{}'.format(lag-1)
                worksheet.write(r+1, c, name)
            else:
                name = '(Yt-Ybar)*(Yt-{}-Ybar)'.format(lag-1)
            worksheet.write(r+1, c, name)
        else:
            if c % 2 != 0:
                worksheet.write(r+1, c, Ylag[r])
            else:
                try:
                    y = (Yt[r]-Ybar)*(Ylag[r]-Ybar)
                except Exception as TypeError:
                    y = 0
                worksheet.write(r+1, c, y)
                r_n = y + r_n
    if c % 2 == 0:
        ac = r_n/r_d
        #print ac
        worksheet.write(r + 2, c, ac)
        ac_list.append(ac)

    end_column = c

row = 0
end_column = end_column+1
worksheet.write(row, end_column, 'Auto-correlation')
for ac in ac_list:
    row = row+1
    worksheet.write(row, end_column, ac)

print ac_list
chart1 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
# [sheetname, first_row, first_col, last_row, last_col].
chart1.add_series({
    'values': [data_set, 1, end_column, row, end_column],
})
worksheet.insert_chart('E1', chart1)

workbook.close()