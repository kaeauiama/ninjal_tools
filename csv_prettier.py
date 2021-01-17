# version 1.0
# last modified 2020-05-14
# --------------------------------------------------------------
# CSV の空行空列を検知して削除するだけのツール
# DB にインサートするときに邪魔になるので作ったもので、現在は用途なし
# --------------------------------------------------------------


import glob, os

def load_csvfile_to_array (abspath):
	""" open the csv and format into 2D array """
	with open(abspath, mode='r') as r:
		arr = []
		line = r.readline()
		while line:
			arr.append(line.rstrip(os.linesep).split(','))
			line = r.readline()
	return arr

def save_array_to_csvfile (arr, abspath):
	""" format the 2D array into csv and save it in the same folder with the abspath """
	csv = ''
	for raw in arr:
		csv += ','.join(raw) + '\n'
	with open(abspath.split('.')[0]+'_new.csv', mode='w') as w:
		w.write(csv)

def remove_last_empty_row (arr):
	""" remove the undermost empty row(s) """
	popped_row = arr.pop()
	while ''.join(popped_row) == '':
		popped_row = arr.pop()
	arr.append(popped_row)
	return arr

def remove_last_empty_column (arr):
	""" remove the rightmost empty column(s) using transposed matrix """
	arr_t = [list(x) for x in zip(*arr)]
	arr_t = remove_last_empty_row (arr_t)
	return [list(x) for x in zip(*arr_t)]

if __name__ == "__main__":
	csv_collection = glob.glob('*.csv')
	for csv_path in csv_collection:
		csv_abspath = os.path.abspath(csv_path)
		csv_arr = load_csvfile_to_array (csv_abspath)
		csv_arr = remove_last_empty_row (csv_arr)
		csv_arr = remove_last_empty_column (csv_arr)
		save_array_to_csvfile (csv_arr, csv_abspath)
	print('処理が終了しました')