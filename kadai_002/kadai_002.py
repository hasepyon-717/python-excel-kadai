import pandas as pd

# データフレームの作成
df = pd.DataFrame({
    '日付':
      ['2023-05-17', '2023-05-18', '2023-05-19', '2023-05-20', '2023-05-21'],
    '社員名': ['山田', '佐藤', '鈴木', '田中', '髙橋'],
    '売上': [100, 200, 150, 300, 250],
    '部門': ['メーカー', '代理店', 'メーカー', '商社', '代理店'],
})

# 売上の平均を求めて新しい列を作成
df['平均売上'] = df['売上'].mean()

# 業績ランクを求める関数「performance」を定義
def performance(rank):
  result = '';
  if rank >= df['売上'].mean() + 50:
    result = 'A';
  elif rank >= df['売上'].mean():
    result = 'B';
  else:
    result = 'C';
  return result

# 「業績ランク」列を作成し、関数「performance」を適用して値を設定
df['業績ランク'] = df['売上'].apply(performance)

# Excelファイルを作成
writer = pd.ExcelWriter('業績.xlsx')
 
# DataFrameオブジェクトをExcelファイルに書き込む
df.to_excel(writer, sheet_name='Sheet1', index=False)

# Excelファイルを閉じる
writer.close()

