#openpyxl

安裝請參考[Openpyxl](https://openpyxl.readthedocs.io/en/default/)

##Workbook
xlsx經過讀取或者新建之後，於此package內的class名稱

###建立新的xlsx物件(Workbook)
'''
from openpyxl import Workbook
wb = Workbook()
'''

###讀取xlsx檔案
'''
from openpyxl import load_workbook
wb = load_workbook('filename.xlsx')
'''

###Workbook的一些操作
'''
from openpyxl import Workbook
wb = Workbook()

#取得資料表(Worksheet)物件
ws = wb.active
#命名資料表物件
ws.title = 'worksheet1'

#取得Workbook內所有worksheet名稱的list
sheetname = wb.sheetnames[0]

#取得指定名稱的資料表
ws = wb[sheetname]

#儲存剛剛建立的xlsx(Workbook)物件
wb.save('destination.xlsx')
’‘’

