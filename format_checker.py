# version 1.0
# last-modified 2020-06-19
# ------------------------------------------------------------------------
# COJADS の CSV のフォーマットをチェックする
# csv_prettier.py にかけて空行空列を削除してから使うこと
# 非破壊的なので指摘をもとに手動でデータを確認・修正することになる
#     ヘッダーが適切な文字列・順序かどうか　空白のセルがないか
#     ID が正しい順序になっているか　県名が正しい表記か　半角英数は残っていないか
#     などを各 CSV ごとにチェックしてコンソールに表示する
# GAS 版「修正支援ツール」と「全体体裁チェック」を統合して洗練した感じのツール
# ------------------------------------------------------------------------


import argparse, glob, os, re


def load_csvfile_to_array (abspath, isUtf8):
	""" load csv from the filepath and turn it into 2D array
	Args:
		abspath (str): the absolute path of the csv file
		isUtf8 (boolean): whether the character code of CSVs is utf8 or not
	Returns:
		arr (arr[][]): the result 2D array representing the original csv
	"""
	if isUtf8 == True:
		char = "utf-8"
	else:
		char = "cp932"
	with open(abspath, mode='r', encoding=char) as r:
		arr = []
		line = r.readline()
		while line:
			arr.append(line.rstrip(os.linesep).split(','))
			line = r.readline()
	return arr


def id_checker(arr):
	arr.pop(0) # タイトル行をパージ
	id_check_result = []

	for i, line in enumerate(arr):
		if str(i+1) != str(line[6]):
			id_check_result.append(str(i+1) + "行目の ID が" + line[6] + "です。")
	
	return id_check_result


def format_checker (arr, isNonmeta):
	""" check if the format of the arr (=csv) match the regulation of COJADS raw data
	Args:
		arr (arr[][]): the 2D arr data representing the original csv file
	Returns:
		correction (list): the notice list for the current csv
	"""
	correction = []
	correct_header = ["xmin","xmax","県","地点","file番号","地点ID","ID","話者","方言テキスト","標準語テキスト","データ名","収録年月日","収録場所","編集担当者","方言チェック担当者","話者生年","話者年齢","話者性別","話題","収録担当者","談話ジャンル","注1番号","注1語形","注1注釈","注2番号","注2語形","注2注釈","注3番号","注3語形","注3注釈"]
	target_header = arr[0]
	excess_item = list(set(target_header)-set(correct_header))
	defici_item = list(set(correct_header)-set(target_header))
	if len(target_header) != len(correct_header):
		correction.append("タイトル行の長さが違います。")
	if excess_item != []:
		correction.append('タイトル行に不要な要素「' + ','.join(excess_item) + "」が含まれています。")
	if defici_item != []:
		correction.append('タイトル行に必要な要素「' + ','.join(defici_item) + "」が含まれていません。")
	if set(target_header) == set(correct_header) and target_header != correct_header:
		correction.append('タイトル行の要素の順序が違うようです。要チェック')
		
	if isNonmeta == False:
		for i, row in enumerate(arr):
			for j, cell in enumerate(row):
				if cell == "" and j < 20:
					correction.append( str(i) + '行' + str(j) + '列が空欄です。')
	
	correction.extend(id_checker(arr))

	return correction


def ken_checker(ken):
	valid_ken_list = ["北海道","青森","岩手","宮城","秋田","山形","福島","茨城","栃木","群馬","埼玉","千葉","東京","神奈川","新潟","富山","石川","福井","山梨","長野","岐阜","静岡","愛知","三重","滋賀","京都","大阪","兵庫","奈良","和歌山","鳥取","島根","岡山","広島","山口","徳島","香川","愛媛","高知","福岡","佐賀","長崎","熊本","大分","宮崎","鹿児島","沖縄"]
	ken_list_with_ken = ["青森県","岩手県","宮城県","秋田県","山形県","福島県","茨城県","栃木県","群馬県","埼玉県","千葉県","神奈川県","新潟県","富山県","石川県","福井県","山梨県","長野県","岐阜県","静岡県","愛知県","三重県","滋賀県","兵庫県","奈良県","和歌山県","鳥取県","島根県","岡山県","広島県","山口県","徳島県","香川県","愛媛県","高知県","福岡県","佐賀県","長崎県","熊本県","大分県","宮崎県","鹿児島県","沖縄県"]

	if ken in valid_ken_list:
		return None
	if ken in ken_list_with_ken:
		str = "「県」列に" + ken + "とありますが、「県」は不要です。"
		return str
	if ken == "東京都":
		str = "「県」列に「東京都」とありますが、「都」は不要です。"
		return str
	if ken == "大阪府" or ken == "京都府":
		str = "「県」列に" + ken + "とありますが、「府」は不要です。"
		return str
	else:
		return "「県」列の文字列が変です。要チェック。"


def textlint(str, redict):
	"""
	チェッカーとしては簡易的なもの、最終チェック用です。
	基本的にはJavaScriptで書いた修正個所チェックを利用してください。
	"""
	correction = []
	str = re.sub(redict["chu"], "", str)
	if redict["kigou"].search(str) != None:
		correction.append("半角記号")
	if redict["ninsho"].search(str) != None:
		correction.append("半角人称")
	if redict["eisu"].search(str) != None:
		correction.append("半角英数")
	return correction


def text_checker (arr):
	"""
	話者記号のチェック
	方言テキストと標準語テキストのチェック
	・半角記号や半角英数字の検知
	・sjisからutf8の変換では理屈上文字化けはかなり生じにくいはずなので検知しません
	"""
	arr.pop(0) # タイトル行をパージ
	correction_all = []
	washa_list = []
	kigou = re.compile(r"[\(\)\[\]\:]")
	ninsho = re.compile(r"([12](SG|PL)|[12](ＳＧ|ＰＬ)|[１２](SG|PL))")
	eisu = re.compile(r"([A-WYZ]|[^XＸ0-9][0-9]|^[0-9])")
	chu = re.compile(r"〔\d+〕")
	privacy = re.compile(r"（Ｒ：(.+?)）")
	redict = {"kigou": kigou, "ninsho": ninsho, "eisu": eisu, "chu": chu, "privacy":privacy}

	for i, line in enumerate(arr):
		washa_list.append(line[7])

		hougen_correction = textlint(line[8], redict)
		if hougen_correction != []:
			correction_all.append("第" + str(i+2) + "行方言　：" + ",".join(hougen_correction) + "（" + line[8] + "）")

		hyojun_correction = textlint(line[9], redict)
		if hyojun_correction != []:
			correction_all.append("第" + str(i+2) + "行標準語：" + ",".join(hyojun_correction) + "（" + line[9] + "）")

		#chu_correction = textlint(line[22] + line[25] + line[28], redict)
		#if chu_correction != []:
		#	correction_all.append("第" + str(i+2) + "行注　　：" + ",".join(chu_correction) + "（" + line[22] + line[25] + line[28] + "）")

	ken_checker_result = ken_checker(arr[0][2])
	if ken_checker_result != None:
		correction_all.append(ken_checker_result)
	print(f'話者は「 {",".join(list(set(washa_list)))} 」の {len(set(washa_list))} 人でした。')
	return correction_all


def argParse():
	argparser = argparse.ArgumentParser()
	argparser.add_argument("-f", "--format", nargs="?", const=True, default=False, help="タイトル行のチェックや空白セル・行・列の検知を行ないます。")
	argparser.add_argument("-t", "--text", nargs="?", const=True, default=False, help="方言テキストや標準語テキスト部分に含まれる半角記号や半角英数字の検知を行ないます（試行版）。")
	argparser.add_argument("-u", "--utf8", nargs="?", const=True, default=False, help="これをつけると UTF8 の CSV を処理できます。")
	argparser.add_argument("-n", "--nonmeta", nargs="?", const=True, default=False, help="メタ情報の部分の処理を飛ばします。")
	return argparser.parse_args()


if __name__ == "__main__":
	args = argParse()

	print("フォルダに含まれるCSVファイルの形式チェックを行ないます。--format オプションをつけるとタイトル行のチェックや空白セル・行・列の検知を行ないます。--text オプションをつけると方言テキストや標準語テキスト部分に含まれる半角記号や半角英数字の検知を行ないます（試行版）。デフォルトでは sjis として開きますが、--utf8 オプションをつけると utf8 として開きます。--nonmeta オプションをつけるとメタ情報の部分の処理を飛ばします。")

	csv_collection = glob.glob('*.csv')

	for csv_path in csv_collection:
		csv_abspath = os.path.abspath(csv_path)
		csv_arr = load_csvfile_to_array (csv_abspath, args.utf8)
		print(csv_path + '：')
		if args.format == True:
			format_check_result = format_checker (csv_arr, args.nonmeta)
			if format_check_result != []:
				print('\n'.join(format_check_result))
		if args.text == True:
			text_check_result = text_checker(csv_arr)
			if text_check_result != []:
				print('\n'.join(text_check_result))

	print('処理が終了しました')