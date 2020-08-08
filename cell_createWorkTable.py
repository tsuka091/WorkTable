#! python3
# coding: utf-8

#Excelの操作に
import openpyxl


#ファイルを開く
# data_only=True 関数のセルも値を読んでくる
book = openpyxl.load_workbook('templete_cell.xlsx', data_only=True)
sheet = book['1']


#セル列定義
# この辺はjsonファイルに定義して使う感じがいいのかな
day_column = 2
start_column = 4
end_column = 5

for i in range(5, 66):
    day = sheet.cell(row=i, column=day_column).value
    if day == '土' or day == '日' or day == None:
        continue
    else:
        sheet.cell(i + 1, start_column, '9:00')
        sheet.cell(i + 1, end_column, '17:30')


#保存
book.save('workTable_cell.xlsx')