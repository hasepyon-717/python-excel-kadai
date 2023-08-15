import openpyxl
# datetimeモジュールのdatetimeクラスをインポートする
from datetime import datetime

# 新規ワークブックを作成
workbook = openpyxl.Workbook()

# アクティブなシートを変数「worksheet」に格納
ws = workbook.active

# 現在の日付（年月日）を取得して、変数todayに代入する
today = datetime.today()

data = [
        ['請求書'],
        ['株式会社ABC','','','','No.','0001'],
        ['〒101-0022 東京都千代田区神田練塀町300','','','','日付',today.strftime('%Y/%m/%d')],
        ['TEL:03-1234-5678 FAX:03-1234-5678'],
        ['担当者名:鈴木一郎 様'],
        ['商品名', '数量', '単価', '金額'],
        ["商品A",2,10000,20000],
        ["商品B",1,15000,15000],
        ["","","","=SUM(E11,E12)"],
        ["合計","","","=E13"],
        ["消費税","","","=E15*0.1"],
        ["税込合計","","","=E15+E16"]
]

for row in data:
    ws.append(row)

ws.insert_cols(1,1)
ws.insert_rows(1,1)
ws.insert_rows(3,1)
ws.insert_rows(8,2)
ws.insert_rows(14,1)

# 日付のtodayオブジェクトを、特定のフォーマットで文字列に変換する
backup_filename = f"請求書_{today.strftime('%Y%m%d')}.xlsx"

workbook.save(backup_filename)
