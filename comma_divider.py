# version 1.0
# last modified 2021-07-27
# --------------------------------------------------------------
# 自作の COJADS DB を作る際に使ったミニツール
# SQLite に挿入するために ,, を , , に変換する
# --------------------------------------------------------------


import glob, os, re

def alternate (path):
	f = open(path, encoding='utf-8')
	data = f.read()  # ファイル終端まで全て読んだデータを返す
	f.close()

	data = re.sub(",,", ", ,", data)

	f = open(path, mode="w", encoding='utf-8')
	f.write(data)
	f.close()

if __name__ == "__main__":
	csv_collection = glob.glob('*.csv')

	for csv_path in csv_collection:
		print(f'${csv_path}の処理を行ないます')
		alternate(csv_path)
		alternate(csv_path)

	print('処理が終了しました')