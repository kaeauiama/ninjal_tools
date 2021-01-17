# version 1.0
# last-modified 2020-03-27
# ------------------------------------------------------------------------
# cojads viewer 用に作ったミニツール。今後使用予定なし
# ------------------------------------------------------------------------

import glob, os, re

if __name__ == "__main__":
	csv_collection = glob.glob('*_reduced.csv')
	for csv_path in csv_collection:
		csv_path_after = csv_path.replace('_reduced', '')
		os.rename(csv_path, csv_path_after)
	print('処理が終了しました')