# version 1.3
# last-modified 2021-07-01
# ------------------------------------------------------------------------
# CSV から .js ファイルを作成します。
# 	data_csv_map.csv -> data_csv.js & data_map.js
# 	data_studies.csv -> data_studies.js
# 	data_events.csv  -> data_events.js
# 解説
# COJADS のホームページの「データDL」「研究業績」「イベント」のデータを更新するときには、
# 	1. 対応する CSV を編集する。
# 	2. この _shaping_data.py を CSV と同じフォルダにおいて、実行する。
# 	3. 作成された .js をすべて data/　フォルダに入れる。
# の順序で作業します。
# ------------------------------------------------------------------------

# -*- coding: utf-8 -*-
import csv, datetime, json, os
from chardet.universaldetector import UniversalDetector


def check_encoding(filepath):
	detector = UniversalDetector()
	with open(filepath, mode='rb') as f:
		for binary in f:
			detector.feed(binary)
			if detector.done:
				break
	detector.close()
	return detector.result['encoding']


def make_datalist (filepath):
	""" read csv from the path and return it in the dict form
	Args:
		filepath (str): the path of the target csv file
	Returns:
		datalist ([dict]): the header row of the csv will be the dict keys
	"""
	datalist = []
	enc = check_encoding(filepath)
	with open (filepath, 'r', encoding=enc) as f:
		for row in csv.DictReader(f):
			datalist.append(row)
	return datalist

def make_categorylist (categoryname, datalist):
	""" make the list of the specified key
	Args:
		categoryname (str): the key string to make out the value list from
		datalist ([dict]): the target database in dict form
	Returns:
		categorylist_simplified (list): the content list with no duplication
	"""
	categorylist_raw = []
	for item in datalist:
		categorylist_raw.append(item[categoryname])
	categorylist_simplified = list(set(categorylist_raw))
	return categorylist_simplified

def outputfile (dict, keyword, filepath):
	""" format the input json into the appropriate .js form to use in the index.html
	Args:
		dict (dict): the dict file to format
		keyword (str): the keyword that is referred to in the index.html
		filepath (str): the original path of the csv
	Returns:
		the output .js string will be saved in the same directory with the input csv
		with the name of 'data_' + keyword + '_' + timestamp + '.js'.
	"""
	# dump（＝ベタ文字列に変換）する
	dumped_str = json.dumps(dict, ensure_ascii=False, indent=2)
	complete_data = "{const data_" + keyword + " = " + dumped_str + ";document.currentScript." + keyword + "dataObj=data_" + keyword + ";}"
	
	# エスケープされる文字を適当に置換する
	#complete_data = complete_data.replace('\\"', '\"')
	#complete_data = complete_data.replace('"[', '[')
	#complete_data = complete_data.replace(']"', ']')
	
	# 出力
	dt_now = datetime.datetime.now()
	dt_now_str = dt_now.strftime('%Y%m%d-%H%M%S')
	with open (os.path.dirname(filepath) + "/data_" + keyword + "_" + dt_now_str + ".js", mode='x', encoding='utf-8') as f:
		f.write(complete_data)
		
def make_data_map_js (filepath):
	datalist = make_datalist(filepath)
	namelist = make_categorylist("name", datalist)
	datalist_full = []
	for item_namelist in namelist:
		name = item_namelist
		linkstr = "<table><tbody>"
		for item_datalist in datalist:
			if item_datalist['name'] == item_namelist:
				coordinates = [ float(item_datalist['coordinate_x']), float(item_datalist['coordinate_y']) ]
				linkstr += r"<tr><td><a href='#' id='csv/" + item_datalist['filename'] + r".csv' onclick='dl(this.id)'>" + item_datalist['filename'] + r"</a></td><td>" + item_datalist['year'] + r"</td><td>" + item_datalist['length'] + "</td></tr>"
		linkstr += "</tbody></table>"
		data_geometry = {"type":"Point", "name": name, "coordinates": coordinates, "link": linkstr}
		data_fragment = {"type": "Feature"}
		data_fragment['geometry'] = data_geometry
		datalist_full.append(data_fragment)
	
	outputfile(datalist_full, "map", filepath)

def make_data_csv_js (filepath):
	datalist = make_datalist(filepath)
	namelist = make_categorylist("name", datalist)
	datalist_all = []
	datalist_all.append({"prefecture": "0", "title": "", "filename": "00", "outline": "<button id='data/cojads_allcsv.zip' onclick='dl(this.id)'>　一括ダウンロード（ZIP）　</button>"})
	for item_of_namelist in namelist:
		name = item_of_namelist
		outlinestr = "<table class='dl'><tbody><tr align='left'><th>ファイル名　　</th><th>収録年</th><th>長さ</th><th>話者数</th><th>公開時期</th><th></th></tr>"
		for item_datalist in datalist:
			if item_datalist['name'] == item_of_namelist:
				prefecture_num = item_datalist['prefecture']
				filename = item_datalist['filename']
				outlinestr += r"<tr><td>" + item_datalist['filename'] + r"　</td><td>" + item_datalist['year'] + r"年　</td><td>" + item_datalist['length'] + "　</td><td>" + item_datalist['speakers'] + "人（男" + item_datalist['male'] + "・女" + item_datalist['female'] + "）</td><td>" + item_datalist['release'].replace('公開', '') + r"</td><td><button id='csv/" + item_datalist['filename'] + r".csv' onclick='dl(this.id)'>　ダウンロード　</button></td></tr>"
		outlinestr += "</tbody></table>"
		datalist_all.append({"prefecture": prefecture_num, "title": name, "filename": filename, "outline": outlinestr})
	
	#県番号順にソートして形を整える
	datalist_ordered = sorted(datalist_all, key = lambda x:x['filename'])
	datalist_full = { "": { "": datalist_ordered } }
	
	outputfile(datalist_full, "csv", filepath)

def make_data_studies_js (filepath):
	datalist = make_datalist(filepath)
	datalist_all = {}
	
	# 年度降順リスト
	yearlist = make_categorylist("year", datalist)
	yearlist.sort(reverse=True)
	
	# 年度内でソートする
	for item_of_yearlist in yearlist:
		year = item_of_yearlist
		datalist_per_year = []
		for item_datalist in datalist:
			data_year = item_datalist['year']
			if data_year == year:
				# 変数の空判定もっとシンプルに書けないのか？ JS の || 演算子みたいに
				data_number    = item_datalist['number']
				data_category  = item_datalist['category']
				data_author    = item_datalist['author']
				data_title     = item_datalist['title']
				data_following = item_datalist['following']
				item_to_append = {"number": data_number, "title": data_author + " " + data_title + "［" + data_category + "］", "outline": data_following}
				if item_datalist['link']:
					item_to_append["href"] = item_datalist['link']
				datalist_per_year.append(item_to_append)
		# ソートする
		datalist_ordered = []
		for i in range(1, len(datalist)+1):
			for each in datalist_per_year:
				if int(each['number']) == i:
					datalist_ordered.append(each)
		datalist_all[str(year)+'年度'] = datalist_ordered
				
	# 形式を整える
	datalist_full = { "": datalist_all }		
	outputfile(datalist_full, "studies", filepath)
	
def make_data_events_js (filepath):
	datalist = make_datalist(filepath)
	datalist_all = {}
	
	# 年度降順リスト
	yearlist = make_categorylist("school-year", datalist)
	yearlist.sort(reverse=True)
	
	# 年度内でソートする
	for item_of_yearlist in yearlist:
		year = item_of_yearlist
		datalist_per_year = []
		for item_datalist in datalist:
			data_year = item_datalist['school-year']
			if data_year == year:
				# 変数の空判定を入れないといけない
				data_number   = item_datalist['number']
				data_title    = item_datalist['title']
				data_evtDate  = item_datalist['evtDate']
				data_evtPlace = item_datalist['evtPlace']
				if item_datalist['link']:
					data_link = item_datalist['link']
				else:
					data_link = ""
				if item_datalist['outline']:
					data_outline = item_datalist['outline']
				else:
					data_outline = ""
				datalist_per_year.append({"number": data_number, "title": data_title, "outline": data_outline, "evtDate": data_evtDate, "evtPlace": data_evtPlace, "href": data_link})
		# ソートする
		datalist_ordered = []
		for i in range(1, len(datalist)+1):
			for each in datalist_per_year:
				if int(each['number']) == i:
					datalist_ordered.append(each)
		datalist_all[str(year)+'年度'] = datalist_ordered
				
	# 形式を整える
	datalist_full = { "": datalist_all }		
	outputfile(datalist_full, "events", filepath)

if __name__ == '__main__':
	# 同一フォルダ内の対象ファイルを取得
	pwd = os.path.dirname(os.path.abspath(__file__))
	filepath_csv_map  = pwd + "/data_csv_map.csv"
	#filepath_news    = pwd + "/data_news.csv"
	filepath_studies  = pwd + "/data_studies.csv"
	filepath_events   = pwd + "/data_events.csv"
	
	# ファイル自由選択（廃止）
	#root = tkinter.Tk()
	#root.withdraw()
	#fTyp = [("","*")]
	#iDir = os.path.abspath(os.path.dirname(__file__))
	#filepath = tkinter.filedialog.askopenfilename(filetypes = fTyp,initialdir = iDir)
	
	# 作成するものだけコメントアウトを外すこと
	#make_data_map_js (filepath_csv_map)
	make_data_csv_js (filepath_csv_map)
	#make_data_news_js (filepath_news)
	#make_data_studies_js (filepath_studies)
	#make_data_events_js (filepath_events)
