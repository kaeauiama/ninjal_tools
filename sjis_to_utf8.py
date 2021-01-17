# version 1.0
# last-modified 2020-05-08
# ------------------------------------------------------------------------
# sjis の CSV を utf8 に変換するだけのプログラム
# utf8 で文字化けになる特殊文字も検出できるとなお良いが……
#
# excel_to_csv.py に統合されたため、今後使用予定なし
# ------------------------------------------------------------------------
#

import glob, os, csv


def load_csvfile_to_array (abspath):
	""" open the csv and convert it into arr
	Args:
		abspath (str): the absolute path of the target csv
	Returns:
		arr (arr): the 2D array that is converted from the input csv
	"""
	with open(abspath, mode='r') as r:
		arr = []
		line = r.readline()
		while line:
			arr.append(line.rstrip(os.linesep).split(','))
			line = r.readline()
	return arr


def save_array_to_csvfile (dataarr, csv_name):
	""" convert 2D array into csv string and save it in the working folder
	Args:
		dataarr (arr): the target array
		csv_name (str): the filename you want to name the result csv string file
	Returns: null
	"""
	csv = ''
	for raw in dataarr:
		csv += ','.join(map(str, raw)) + '\n'
	with open(csv_name + '_utf8.csv', mode='w', encoding='utf-8') as w:
		w.write(csv)


if __name__ == "__main__":
	csv_collection = glob.glob('*.csv')
	
	print(csv_collection)
	
	for each_csv in csv_collection:
		save_array_to_csvfile(load_csvfile_to_array(each_csv), each_csv.replace('.csv', ''))
	
	print('処理が終了しました')
