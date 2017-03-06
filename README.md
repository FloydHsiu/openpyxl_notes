#openpyxl_notes

>安裝請參考[Openpyxl](https://openpyxl.readthedocs.io/en/default/)

隨附的`test.xlsx`及`test2.xlsx`為兩天不同的匯率表
以此兩個xlsx可以練習openpyxl的操作

##Workbook
xlsx經過讀取或者新建之後，於此package內的class名稱

###建立新的xlsx物件(Workbook)
```
from openpyxl import Workbook
wb = Workbook()
```

###讀取xlsx檔案
```
from openpyxl import load_workbook
wb = load_workbook('filename.xlsx')
```

###Workbook的一些操作
```
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
```

##worksheet
worksheet是由很多的Cell物件組成，所以讀取ws內的值會取得Cell，如果要更動Cell的值，後面將會討論到
###worksheet的一些操作
```
ws = wb.active

#讀取ws內'A1'值(會取得單一一個Cell)(有兩種方法)
ws['A1']
ws.cell(row=1, column=1)

#讀取ws內'A'行的值(會取得一個Cell的List)
ws['A']
ws['A':'B']

#讀取ws內'1'列的值(會取的一個Cell的List)
ws['1']
ws['1':'2']

#將每一行當成一個tuple的清單
tuple(ws.columns)

#將每一列當成一個tuple的清單
tuple(ws.rows)
```

##Cell
儲存worksheet的單位

###Cell的一些操作
```
#取讀Cell的值
ws['A1'].value

#更動Cell的值
ws['A1'].value = 20
```

