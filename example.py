import xl
import xlsxwriter
import yaml

with open('data.yml', 'r') as f:
    data = yaml.load(f.read())

with open('rdata.yml', 'r') as f:
    rdata = yaml.load(f.read())

book = xlsxwriter.Workbook('demographics.xlsx')
sheet = book.add_worksheet('Key Metrics')
(x, y) = xl.xlTable(sheet, data, name='rows')
xl.xlTable(sheet, rdata, name='records', coord=(x+1,y+1))
book.close()

