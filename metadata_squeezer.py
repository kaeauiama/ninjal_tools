# version 1.0
# last-modified 2020-05-08
# ------------------------------------------------------------------------
# COJADS の CSV 生データから viewer 用のファイルを作る
# それぞれの csv から 01_b_099.csv => 01_b_099_reduced.csv のようにメタデータ削除した csv をつくる（軽量化）
# csv からメタデータだけ抽出した metadata.js
# csv から話者情報だけ抽出した speakerdata.js
# menubar の表示に使う県名とファイル名の入っただけの menubar.js
#
# prettier, formatter にかけて CSV の形式を整えてから利用すること
# また web 上での処理を考えているため、 csv は utf-8 で保存する仕様にしている
#
# 現在は viewer の更新予定もなく、使用予定なし
# ------------------------------------------------------------------------


import glob, os, csv, json


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
	with open(csv_name + '.csv', mode='w', encoding='utf-8') as w:
		w.write(csv)


def save_array_to_jsfile (dataarr, keyword):
	""" convert 2D array into json-like .js string and save it in the working folder
	Args:
		dataarr (arr): the target array
		keyword (str): the filename you want to name the result .js string file
	Returns: null
	"""
	# json 形式にする
	filenamelist = []
	for row in dataarr:
		if not row[0] == 'filename':
			filenamelist.append(row[0])
	#filenamelist = get_unique_list(filenamelist).sort() ←うまく動かないが上書きされるのでこれは不要
	
	labellist = dataarr[0]
	datajson = {}
	
	for item_filename in filenamelist:
		data_per_file = []
		for row in dataarr:
			if row[0] == item_filename:
				tempdict = {}
				for i, item in enumerate(row):
					tempdict[labellist[i]] = item
				data_per_file.append(tempdict)
		datajson[item_filename] = data_per_file
	
	# dump（＝ベタ文字列に変換）する
	datastr = json.dumps(datajson, ensure_ascii=False, indent=2)
	complete_data = "{const data_" + keyword + " = " + datastr + ";document.currentScript." + keyword + "dataObj=data_" + keyword + ";}"
	
	# エスケープされる文字を適当に置換する
	complete_data = complete_data.replace('\\"', '"')
	complete_data = complete_data.replace('"[', '[')
	complete_data = complete_data.replace(']"', ']')
	
	with open (keyword+'.js', mode='x', encoding='utf-8') as f:
		f.write(complete_data)		


def get_unique_list(arr):
	""" take the list/array, get lid of its duplicated member, return the list/array with no duplication """
	seen = []
	return [x for x in arr if not seen.append(x) and seen.count(x) == 2]


def make_metadata (name, dataarr):
	""" extract the metadata from the target csv array
	Args:
		name (str): the filename of the target csv
		  You should implement this because the original csv does not include its file name info in it
		  the filename is very important because it will be used as the key binding the reduced csv, metadata, and speakerdata
		dataarr (arr): the target 2D array representing csv
	Returns:
		metadata_arr (arr): the metadata extracted
		  Basically it should be consist of the only one record after simplification
	"""
	metadata_arr = []
	for i, raw in enumerate(dataarr):
		# filename, 県, 地点, file番号, 地点ID, データ名, 収録年月日, 収録場所, 編集担当者,　収録担当者,　談話ジャンル, 話題
		if not i == 0:
			metadata_arr.append([name, raw[2], raw[3], raw[4], raw[5], raw[10], raw[11], raw[12], raw[13], raw[18], raw[19], raw[17]])
	metadata_arr = get_unique_list(metadata_arr)
	
	# 話題だけは同一ファイル内で複数の値を持ちうるので、結合する
	metadata_arr = topic_joiner(metadata_arr)
	return metadata_arr


def topic_joiner(arr):
	""" join the multiple topics and reduce the redundant raws """
	# 一度「話題」をはずしたメタデータをつくる
	topic_popped_arr = []
	for row in arr:
		row.pop(-1)
		topic_popped_arr.append(row)
	topic_popped_arr = get_unique_list(topic_popped_arr)
	# 話題を取得して、複数あるようなら結合して追加する
	for row_t in topic_popped_arr:
		topic_list = []
		for row_o in arr:
			if row_t[2] == row_o[2]:
				topic_list.append(row_o[11])
		topic_str = '、'.join(topic_list)
		row_t.append(topic_str)
	return topic_popped_arr


def make_speakerdata (name, dataarr):
	""" extract the speakerdata from the target csv array
	Args:
		name (str): the filename of the target csv
		dataarr (arr): the target 2D array representing csv
	Returns:
		speakerdata_arr (arr): the speakerdata extracted
	"""
	speakerdata_arr = []
	for i, raw in enumerate(dataarr):
		# filename, 話者, 話者生年, 話者年齢, 話者性別
		if not i == 0:
			speakerdata_arr.append([name, raw[7], raw[14], raw[15], raw[16]])
	speakerdata_arr = get_unique_list(speakerdata_arr)
	return speakerdata_arr


def make_and_save_menubardara (dataarr):
	""" make the prefecture (key) -> filenames (values) dict and save in the appropriate form
	Args:
		dataarr (arr): the target 2D array representing csv
	"""
	menubardata = {}
	plist = ["北海道","青森","岩手","宮城","秋田","山形","福島","茨城","栃木","群馬","埼玉","千葉","東京","神奈川","新潟","富山","石川","福井","山梨","長野","岐阜","静岡","愛知","三重","滋賀","京都","大阪","兵庫","奈良","和歌山","鳥取","島根","岡山","広島","山口","徳島","香川","愛媛","高知","福岡","佐賀","長崎","熊本","大分","宮崎","鹿児島","沖縄"]
	for pref in plist:
		list_for_pref = []
		for item in dataarr:
			if pref == item[1]:
				list_for_pref.append(item[0])
		menubardata[pref] = list_for_pref
	
	datastr = json.dumps(menubardata, ensure_ascii=False, indent=2)
	print(datastr)
	complete_data = "{const data_menubar = " + datastr + ";document.currentScript.menubardataObj=data_menubar;}"
	
	# エスケープされる文字を適当に置換する
	complete_data = complete_data.replace('\\"', '"')
	complete_data = complete_data.replace('"[', '[')
	complete_data = complete_data.replace(']"', ']')
	
	with open ('menubardata.js', mode='x', encoding='utf-8') as f:
		f.write(complete_data)		


def reduce_csv_arr (name, arr):
	""" make the simple 2D array by removing the metadata and speakerdata, which are divided into the other datafile
	Args:
		name (str): the filename of the csv
		arr (arr): the target 2D array representing the full csv
	Returns:
		reduced_csv (arr): the 2D array representing the reduced csv
	"""
	reduced_csv = []
	for i, raw in enumerate(arr):
		# filename, xmin, xman, ID, 話者, 方言テキスト, 標準語テキスト
		# 現状、注釈は無視する
		if i == 0:
			reduced_csv.append(['filename' , raw[0], raw[1], raw[6], raw[7], raw[8], raw[9]])
		else:
			reduced_csv.append([name, raw[0], raw[1], raw[6], raw[7], raw[8], raw[9]])
	return reduced_csv


if __name__ == "__main__":
	csv_collection = glob.glob('*.csv')
	metadata_arr = [['filename', 'prefecture', 'place', 'filenumber', 'placeid', 'data', 'recordeddate', 'recordedplace', 'editor', 'topic', 'recorder', 'genre']]
	speakerdata_arr = [['filename', 'speaker', 'birthyear', 'age', 'sex']]
	
	print(csv_collection)
	
	# 各 CSV に対して、metadata, speakerdata への追加および reducedcsv の作成を行う
	for csv_path in csv_collection:
		csv_filename = csv_path.replace('.csv', '')
		csv_arr = load_csvfile_to_array(os.path.abspath(csv_path))
		metadata_arr += make_metadata(csv_filename, csv_arr)
		speakerdata_arr += make_speakerdata(csv_filename, csv_arr)
		reducedcsv_arr = reduce_csv_arr(csv_filename, csv_arr)
		#save_array_to_csvfile(reducedcsv_arr, csv_path.replace('.csv', '') + '_reduced')
		
	# metadata, speakerdata を書き出す
	# 不要 save_array_to_csvfile(metadata_arr, 'metadata')
	# 不要 save_array_to_csvfile(speakerdata_arr, 'speakerdata')
	#save_array_to_jsfile(metadata_arr, 'metadata')
	#save_array_to_jsfile(speakerdata_arr, 'speakerdata')
	make_and_save_menubardara(metadata_arr)
	
	print('処理が終了しました')