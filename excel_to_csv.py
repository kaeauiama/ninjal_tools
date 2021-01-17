# version 3.0
# last-modified 2021-01-16
# --------------------------------------------------------------
# 校正作業の終了した Excel からコメント列等を削除し、
# それを COJADS サイト上で配布する CSV に変換し、ついでに ZIP を作成します
# デフォルトでは SJIS と UTF8 の CSV を作成するようになっています
# --sjis オプションをつけると sjis の CSV のみを作成します
# --utf8 オプションをつけると utf8 の CSV のみを作成します
# --------------------------------------------------------------
# TODO: 不要列削除の機能はあとになって追加したので Excel の読み込みが２回行なわれている

# -*- coding: utf-8 -*-
import argparse, glob, xlrd, pandas, sys, re, io, os, zipfile


def remove_unnecessary_column (path):
    df = pandas.read_excel(path, dtype="object")
    expected_header = ["xmin","xmax","県","地点","file番号","地点ID","ID","話者","方言テキスト","標準語テキスト","データ名","収録年月日","収録場所","編集担当者","方言チェック担当者","話者生年","話者年齢","話者性別","話題","収録担当者","談話ジャンル","注1番号","注1語形","注1注釈","注2番号","注2語形","注2注釈","注3番号","注3語形","注3注釈"]
    actual_header = df.columns.values.tolist()
    for x in actual_header:
        if x not in expected_header:
            df=df.drop(columns=x)

    # 古いファイルを _old に改称して新しいのを保存
    old_path = path.replace(".xls", "_old.xls")
    os.rename(path, old_path)
    df.to_excel(path, index=False)


def load_excelfile_to_array (path):
	sheet = xlrd.open_workbook(path).sheets()[0]
	arr = []
	for r in range(0, sheet.nrows):
		line = []
		for c in range(0, sheet.ncols):
			if c == 6:
				line.append(str(sheet.cell(r,c).value).replace('\n','').replace('.0',''))
			else:
				line.append(str(sheet.cell(r,c).value).replace('\n',''))
		arr.append(line)
	return arr


def save_array_to_csvfile (dataarr, csv_name, isOnlySjis, isOnlyUtf8):
	csv = ''
	for raw in dataarr:
		csv += ','.join(map(str, raw)) + '\n'

	# --utf8 指定のときは飛ばす
	if isOnlyUtf8 == False:
		with open(csv_name + '.csv', mode='w', encoding='cp932') as w:
			w.write(csv)

	# --sjis 指定のときは飛ばす
	if isOnlySjis == False:
		with open(csv_name + '_utf8.csv', mode='w', encoding='utf-8') as w:
			w.write(csv)


def excel_to_csv (sjis, utf8):
	excel_collection = glob.glob('*.xlsx') + glob.glob('*.xls')

	# 処理を行なうファイルを表示
	print(excel_collection)

	err_count = 0

	# 不要列削除 -> 読み込み -> CSV として保存
	for excel_path in excel_collection:
		excel_abspath = os.path.abspath(excel_path)
		try:
			remove_unnecessary_column (excel_abspath)
		except:
			print('columnのトリム処理失敗：', excel_path)
			err_count += 1
			continue
		try:
			excel_arr = load_excelfile_to_array (excel_abspath)
		except:
			print('arrへの変換失敗：', excel_path)
			err_count += 1
			continue
		try:
			csv_name = re.sub(r'\.xlsx|\.xls','',excel_path)
			save_array_to_csvfile (excel_arr, csv_name, sjis, utf8)
		except:
			print('csvへの変換失敗：', excel_path)
			err_count += 1
			continue
		print('変換終了：', excel_path)
	
	return err_count


def make_zip (isOnlyUtf8, isOnlySjis):
	csv_sjis_collection = sorted(list(set(glob.glob('*.csv')) - set(glob.glob('*_utf8.csv'))))
	csv_utf8_collection = sorted(glob.glob('*_utf8.csv'))

	# --utf8 指定のときは飛ばす
	if isOnlyUtf8 == False:
		#print(csv_sjis_collection)
		z = zipfile.ZipFile("allcsv.zip", "w")
		for i in csv_sjis_collection:
			z.write(i)
		z.close()
		print('SJIS圧縮終了：allcsv.zip')

	# --sjis 指定のときは飛ばす
	if isOnlySjis == False:
		#print(csv_utf8_collection)
		z = zipfile.ZipFile("allcsv_utf8.zip", "w")
		for i in csv_utf8_collection:
			z.write(i)
		z.close()
		print('UTF8圧縮終了：allcsv_utf8.zip')


def argParse():
	argparser = argparse.ArgumentParser()
	argparser.add_argument("-s", "--sjis", nargs="?", const=True, default=False, help="SJIS のみ作成")
	argparser.add_argument("-u", "--utf8", nargs="?", const=True, default=False, help="UTF8 のみ作成")
	return argparser.parse_args()


if __name__ == '__main__':
	args = argParse()

	if args.sjis == True and args.utf8 == True:
		print("オプションが矛盾しています。")
		exit(1)
	
	# Excel を CSV に変換
	# エラーカウントを返却
	err = excel_to_csv(args.sjis, args.utf8)
	
	# CSV から ZIP を作成
	# ただしエラーが出ているときはスキップ
	if (err == 0):
		print('続けて一括ダウンロード用の ZIP を作成しますか？ しない場合は n を入力：', end='')
		str = input()
		if (str != "n"):
			try:
				make_zip (args.sjis, args.utf8)
			except:
				print("ZIP の作成に失敗しました。")
	else:
		print('変換失敗したデータがあるため、ZIP 作成は行なえません。')
		
	print('処理が終了しました')
