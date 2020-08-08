#! python3
# coding: utf-8

#Excelの操作に
import openpyxl
#ファイルのコピーに
import shutil

#ファイルを開く
# data_only=True 関数のセルも値を読んでくる
book = openpyxl.load_workbook('templete_cell.xlsx', data_only=True)
sheet = book['1']


#セル列定義
# この辺はjsonファイルに定義して使う感じがいいのかな
day_cell_num = 1
start_cell_num = 3
end_cell_num = 4

for row in ws.iter_rows(min_row=5, max_row=35):
    day = row[day_cell_num].value
    if day == '土' or day == '日':
        row[start_cell_num].value = ''
        row[end_cell_num].value = ''
    else:
        row[start_cell_num].value = '9:00'
        row[end_cell_num].value = '17:30'        


#保存
book.save('workTable.xlsx')



# #ワークブック新規作成
# book = openpyxl.Workbook()

# #シート名変更
# sheet = book.active
# sheet.title = 'sample'

# #範囲を指定してセルを取得する
# cells = sheet['A1':'B3']
# i = 0
# for row in cells:
#     for cell in row:
#         cell.value = i
#         i += 1

# #ワークブックに名前を付けて保存
# book.save('sample.xlsx')