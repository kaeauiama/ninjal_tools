# version 3.2
# last-modified 2020-05-01
# ------------------------------------------------------------------------
# TextGrid から Excel を作る
# -s (--serial)	処理終了後、次のTextGrid選択画面を開きます
# -a (--all)	選択されたフォルダ内のTextGridを全て処理します
# バグフィックスした場合は Laravel アプリのほうにも反映させること
# ------------------------------------------------------------------------


# -*- coding: utf-8 -*-
import argparse, openpyxl, glob, sys, re, os, tkinter, tkinter.filedialog
from chardet.universaldetector import UniversalDetector


def check_encoding(file_path):
	detector = UniversalDetector()
	with open(file_path, mode='rb') as f:
		for binary in f:
			detector.feed(binary)
			if detector.done:
				break
	detector.close()
	return detector.result['encoding']


def make_excelfile(filepath):
	#文字コードがutf-8ならShift-JISに変換する；Shift-JISならそのまま
	charset = check_encoding(filepath)
	if charset != 'SHIFT_JIS':
		filedata = open(filepath, encoding=charset)
		f = filedata.readlines()
		filedata.close()
		for row in f:
			row.encode('shift_jis')
	else:
		filedata = open(filepath, encoding='shift_jis')
		f = filedata.readlines()
		filedata.close()
	
	#出力ファイル（.xlsx）を作る
	wb = openpyxl.Workbook()
	sheet = wb.active
	output_file_path = re.sub('.TextGrid', "_fromTG.xlsx", filepath)
	
	#TextGridを１行ごとに処理する
	#まずはタイトル行をつくる
	sheet.cell(row=1, column=1, value='xmin')
	sheet.cell(row=1, column=2, value='xmax')
	sheet.cell(row=1, column=3, value='tier')
	sheet.cell(row=1, column=4, value='text')
	
	#変数の初期化
	xmin = 0
	xmax = 0
	tier = ''
	text = ''
	numb = 0
	mark = ''
	line = 1
	
	#順次内容行をつくる
	for i, row in enumerate(f):
		try:
			# item が来たら tier=i+2行目の name = "___" の部分を代入
			if re.match(r'.*item \[\d.*', f[i]):
				tier = re.match(r'.*name = (.*?)$', f[i+2]).group(1).replace('"', '', 2)

			# intervals が来たら xmin=i+1行目, xmax=i+2行目, tier=据え置き, text=i+3行目
			elif 'intervals [' in row:
				xmin = re.match(r'.*xmin = ([\d\.].*)$', f[i+1]).group(1)
				xmax = re.match(r'.*xmax = ([\d\.].*)$', f[i+2]).group(1)
				# text には余計な改行が入っていることがあるので場合分けする
				if len(f) > i+4 and not 'intervals' in f[i+4] and not 'item' in f[i+4]:
					text = re.match(r'.*text = (.*?)$', f[i+3]).group(1)
					text += f[i+4]
				else:
					text = re.match(r'.*text = (.*?)$', f[i+3]).group(1)

				# text が空欄のときは無視する
				text = re.sub(r'[\"\n]', '', text)
				if text.replace(' ', '') == '':
					continue

				line = line + 1
				sheet.cell(row=line, column=1, value=float(xmin))
				sheet.cell(row=line, column=2, value=float(xmax))
				sheet.cell(row=line, column=3, value=tier)
				sheet.cell(row=line, column=4, value=text)
			
			# points が来たら numb=i+1行目, mark=i+2行目, tier=据え置き
			elif 'points [' in row:
				numb = re.match(r'.*number = ([\d\.].*)$', f[i+1]).group(1)
				# mark には余計な改行が入っていることがあるので場合分けする
				if len(f) > i+3 and not 'points' in f[i+3] and not 'item' in f[i+3]:
					mark = re.match(r'.*mark = (.*?)$', f[i+2]).group(1)
					mark += f[i+3]
				else:
					mark = re.match(r'.*mark = (.*?)$', f[i+2]).group(1)

				mark = re.sub(r'[\"\n]', '', mark)
				if mark.replace(' ', '') == '':
					continue

				line = line + 1
				sheet.cell(row=line, column=1, value=float(numb))
				sheet.cell(row=line, column=3, value=tier)
				sheet.cell(row=line, column=4, value=mark)
		except Exception as ex:
			print(str(i+1) + '行目でエラー')
			print(ex)

	#保存する
	wb.save(output_file_path)
	print(f'{os.path.basename(output_file_path)} を作成しました。')


def argParse():
	argparser = argparse.ArgumentParser()
	argparser.add_argument("-s", "--serial", nargs="?", const=True, default=False, help="処理終了後、次のTextGrid選択画面を開きます")
	argparser.add_argument("-a", "--all", nargs="?", const=True, default=False, help="選択されたフォルダ内のTextGridを全て処理します")
	return argparser.parse_args()


if __name__ == '__main__':
	args = argParse()

	print("TextGridファイルを選択してください。「（元のファイル名_fromTG）」という名前のExcelファイルが作成されます。\n--serialオプションをつけると連続で処理できます。--all オプションをつけるとフォルダ選択ができ、フォルダ内のTextGrid全てを一気に処理します。")

	if args.all == False:
		# ダイアログを開いてTextGridファイルを選択
		# 連続処理の場合に前回のフォルダ位置を保持するために変なコードになっている＠改善したい
		root = tkinter.Tk()
		root.withdraw()
		fTyp = [("テキストグリッド","*.TextGrid")]
		iDir = os.path.abspath(os.path.dirname(__file__))
		filepath = tkinter.filedialog.askopenfilename(filetypes = fTyp,initialdir = iDir)
		while True:
			if filepath=="":
				print("キャンセルされました。")
				sys.exit(1)
			make_excelfile(filepath)
			if args.serial == False:
				break
			root = tkinter.Tk()
			root.withdraw()
			fTyp = [("テキストグリッド","*.TextGrid")]
			filepath = tkinter.filedialog.askopenfilename(filetypes = fTyp, initialdir = os.path.abspath(os.path.dirname(filepath)))
	
	if args.all == True:
		# ダイアログを開いてフォルダを選択
		root = tkinter.Tk()
		root.withdraw()
		iDir = os.path.abspath(os.path.dirname(__file__))
		folderpath = tkinter.filedialog.askdirectory(initialdir = iDir)
		textgridlist = glob.glob(folderpath + "/*.TextGrid")
		if textgridlist == []:
			print("TextGridを発見できませんでした。")
			sys.exit(1)
		for item in textgridlist:
			make_excelfile(item)
	
	print("処理が終了しました。")

