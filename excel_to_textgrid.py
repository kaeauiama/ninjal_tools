# version 4.0
# last-modified 2021-01-30
# ------------------------------------------------------------------------
# Excel を TextGrid へと変換する
# TextGrid にない情報として音声の総時間が必要なので音声ファイルも読み込む
# バグフィックスした場合は Laravel アプリのほうにも反映させること
# ------------------------------------------------------------------------

# -*- coding: utf-8 -*-
import argparse, codecs, xlrd, sys, re, io, os, tkinter, tkinter.filedialog, wave, glob, logging


# ログ出力用の設定
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
f1 = logging.Formatter('line %(lineno)d: [%(levelname)s] %(message)s')
stdout = logging.StreamHandler()
stdout.setLevel(logging.ERROR)
stdout.setFormatter(f1)
logger.addHandler(stdout)
f2 = logging.Formatter('%(asctime)s [%(filename)s: %(lineno)d] [%(levelname)s] %(message)s')
fileout = logging.FileHandler(filename="_excel_to_textgrid.log")
fileout.setLevel(logging.DEBUG)
fileout.setFormatter(f2)
logger.addHandler(fileout)
logger.propagate = False

# 変数を出力するためのヘルパー
def logvar(namelist, varlist):
	loglist = []
	for i in range(0, len(namelist)):
		if varlist[i] is not None:
			loglist.append(f'{namelist[i]} = "{varlist[i]}"')
	logstr = ', '.join(loglist)
	logger.info(f'変数：{logstr}')
	return None


#話者リストをつくる関数
def split_speaker(speaker):
	#全角半角を統一する、ここでは全角に
	full_width = 'ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ'
	half_width = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
	speakers_list = []
	isAlphabet = True
	speaker = re.sub(r'[\'\s]', '', speaker)
	for i in range(len(speaker)):
		if speaker[i] in full_width:
			speakers_list.append(speaker[i])
		elif speaker[i] in half_width:
			speakers_list.append(chr(ord(speaker[i]) - ord('A') + ord('Ａ')))
		else:
			isAlphabet = False
	if isAlphabet == False:
		speakers_list.append(speaker)
	return speakers_list



def get_speaker_xmin_xmax_list(sheet):
	xmin_col = sheet.col(0)
	xmax_col = sheet.col(1)
	speaker_col = sheet.col(7)
	text_col = sheet.col(8)

	#エラー判定
	if not (len(xmin_col) == len(xmax_col) and len(xmax_col) == len(speaker_col)):
		print('xmin, xmax, 話者の行数が揃っていないので揃えてください')
		logvar(
			["xmin_col", "xmax_col", "speaker_col"],
			[len(xmin_col), len(xmax_col), len(speaker_col)]
		)
		raise Exception

	#エクセルファイルから話者、xmin、xmax、方言テキストを抜き出したリストをつくり、話者でソートする
	try:
		speaker_xmin_xmax_list = []
		for i in range(1, len(xmin_col)):
			if not 'empty' in str(speaker_col[i]):
				speakers = re.sub('text:', '', str(speaker_col[i])) # text:'A' のように	typeとともに入っているからトリミングする
				speakers_list = split_speaker(speakers)
				xmin = re.sub('number:', '', str(xmin_col[i]))
				xmax = re.sub('number:', '', str(xmax_col[i]))
				text = re.sub('text:', '', str(text_col[i]))
				text = re.sub(r'〔\d+?〕', '', text)
				text = re.sub('\\\\u3000', '　', text) #全角スペースがエスケープされているので戻す
				text = re.sub('\'', '"', text)
				for j in range(len(speakers_list)):
					speaker_xmin_xmax_list.append((speakers_list[j], float(xmin), float(xmax), text))
	except Exception as e:
		# 変数のブロック大丈夫？
		logger.exception(f'{e}')
		logvar(
			["i", "j", "xmin", "xmax"],
			[i, j, xmin, xmax]
		)
	speaker_xmin_xmax_list.sort()
	logger.info('正常終了：get_speaker_xmin_xmax_list')	
	return speaker_xmin_xmax_list


#話者数を調べる関数
def get_speaker_size(speaker_xmin_xmax_list):
	size = 0
	checked_speaker = []
	for speaker, xmin, xmax, text in speaker_xmin_xmax_list:
		if speaker in checked_speaker:
			continue
		size += 1
		checked_speaker.append(speaker)

	logger.info('正常終了：get_speaker_size')
	return size


#話者ごとに配列を分けて、それをsplit_by_speaker_listにまとめる関数
#split_by_speaker_list[i]に各話者の発話データ（xmin, xmax, text）が入る
def split_by_speaker(speaker_xmin_xmax_list):
	list_split_by_speaker = []
	now_speaker = speaker_xmin_xmax_list[0][0]
	speaker_content = []
	for i in range(len(speaker_xmin_xmax_list)):
		if speaker_xmin_xmax_list[i][0] == now_speaker:
			speaker_content.append(speaker_xmin_xmax_list[i])
		else:
			list_split_by_speaker.append(speaker_content)
			speaker_content = []
			speaker_content.append(speaker_xmin_xmax_list[i])
			now_speaker = speaker_xmin_xmax_list[i][0]
	list_split_by_speaker.append(speaker_content)

	logger.info('正常終了：split_by_speaker')
	return list_split_by_speaker


#TextGridを書き出す関数
def compose_textgrid(xmax, list_split_by_speaker, speaker_size):
	try:
		data = ""
		data += 'File type = "ooTextFile"\n'
		data += 'Object class = "TextGrid"\n'
		data += '\n'
		data += 'xmin = 0\n'
		data += 'xmax = {}\n'.format(xmax)
		data += 'tiers? <exists>\n'
		data += 'size = {}\n'.format(speaker_size)
		data += 'item []:\n'
		for i in range(speaker_size):
			data += '    item [{}]:\n'.format(i + 1)
			data += '        class = "IntervalTier"\n'
			data += '        name = "{}"\n'.format(list_split_by_speaker[i][0][0])
			data += '        xmin = 0\n'
			data += '        xmax = {}\n'.format(xmax)
			data += '        intervals: size = {}\n'.format(len(list_split_by_speaker[i]))
			for j in range(len(list_split_by_speaker[i])):
				data += '        intervals [{}]:\n'.format(j + 1)
				data += '            xmin = {}\n'.format(list_split_by_speaker[i][j][1])
				data += '            xmax = {}\n'.format(list_split_by_speaker[i][j][2])
				data += '            text = {}\n'.format(list_split_by_speaker[i][j][3])
	except Exception as e:
		logger.exception(f'{e}')
		logvar(["i", "j"], [i, j])
		raise Exception
	logger.info('正常終了：compose_textgrid')
	return data


# 話者ごとにソートした際に第n行のxmaxが第n+1行のxminを越えて（＝発話区間が被って）しまうことがあるが、
# その場合に第n行のxmaxを狭めて発話区間の被りをなかったことにする
def change_split_by_speaker_list(split_by_speaker_list, speaker_size):
	for i in range(speaker_size):
		for j in range(len(split_by_speaker_list[i]) - 1):
			if split_by_speaker_list[i][j][2] > split_by_speaker_list[i][j + 1][1]:
				speaker = split_by_speaker_list[i][j][0]
				xmin = split_by_speaker_list[i][j][1]
				xmax = split_by_speaker_list[i][j + 1][1]
				text = split_by_speaker_list[i][j][3]
				split_by_speaker_list[i][j] = (speaker, xmin, xmax, text)
	
	logger.info('正常終了：change_split_by_speaker_list')
	return split_by_speaker_list


#TextGridに空白区間を追加する関数
def add_silent_section(split_by_speaker_list, speaker_size, entire_xmax):
	split_by_speaker_list2 = []
	for i in range(speaker_size):
		tmp = []
		if split_by_speaker_list[i][0][1] != 0:
			speaker = split_by_speaker_list[i][0][0]
			xmin = 0
			xmax = split_by_speaker_list[i][0][1]
			text = '""'
			tmp.append((speaker, xmin, xmax, text))
		for j in range(len(split_by_speaker_list[i]) - 1):
			speaker = split_by_speaker_list[i][j][0]
			xmin = split_by_speaker_list[i][j][1]
			xmax = split_by_speaker_list[i][j][2]
			text = split_by_speaker_list[i][j][3]
			tmp.append((speaker, xmin, xmax, text))
			if split_by_speaker_list[i][j][2] != split_by_speaker_list[i][j + 1][1]:
				speaker = split_by_speaker_list[i][j][0]
				xmin = split_by_speaker_list[i][j][2]
				xmax = split_by_speaker_list[i][j + 1][1]
				text = '""'
				tmp.append((speaker, xmin, xmax, text))
		speaker = split_by_speaker_list[i][len(split_by_speaker_list[i]) - 1][0]
		xmin = split_by_speaker_list[i][len(split_by_speaker_list[i]) - 1][1]
		xmax = split_by_speaker_list[i][len(split_by_speaker_list[i]) - 1][2]
		text = split_by_speaker_list[i][len(split_by_speaker_list[i]) - 1][3]
		tmp.append((speaker, xmin, xmax, text))
		if split_by_speaker_list[i][len(split_by_speaker_list[i]) - 1][2] < entire_xmax:
			speaker = split_by_speaker_list[i][len(split_by_speaker_list[i]) - 1][0]
			xmin = split_by_speaker_list[i][len(split_by_speaker_list[i]) - 1][2]
			xmax = entire_xmax
			text = '""'
			tmp.append((speaker, xmin, xmax, text))
		split_by_speaker_list2.append(tmp)

	logger.info('正常終了：add_silent_section')
	return split_by_speaker_list2


def argParse():
	argparser = argparse.ArgumentParser()
	argparser.add_argument("-s", "--serial", nargs="?", const=True, default=False, help="処理終了後、次のファイル選択画面を開きます")
	argparser.add_argument("-n", "--noaudio", nargs="?", const=True, default=False, help="音声ファイルの長さを参照せずにTextGridを作ります（非推奨）")
	argparser.add_argument("-a", "--all", nargs="?", const=True, default=False, help="フォルダ内のExcelを一括変換します（自動的に --noaudio オプションがつきます）。")
	return argparser.parse_args()


def get_excel_path (pwdpath=None):
	root = tkinter.Tk()
	root.withdraw()
	fTyp = [("エクセルファイル","*.xlsx;*.xls;*.xlsm")]

	# serial の場合は前回処理したファイルのディレクトリで選択画面を開く
	iDir = pwdpath if pwdpath is not None else os.path.abspath(os.path.dirname(__file__))
	item = tkinter.filedialog.askopenfilename(filetypes = fTyp,initialdir = iDir)

	if item == "":
		print("キャンセルされました。")
		logger.info('ファイル選択キャンセルにより終了')
		exit(1)

	return item


def get_excel_collection ():
	excel_collection = sorted(glob.glob('*.xls*'))
	return excel_collection


def get_xmax_from_audio (item):
	#同名のwavファイルがあるなら自動的に開く
	wavpath = re.sub('.xls$|.xlsx$|.xlsm$', ".wav", item)
	if not os.path.exists(wavpath):
		root = tkinter.Tk()
		root.withdraw()
		fTyp = [("音声ファイル","*.wav")]
		iDir = os.path.abspath(os.path.dirname(item))
		wavpath = tkinter.filedialog.askopenfilename(filetypes = fTyp,initialdir = iDir)
	if wavpath=="":
		print("キャンセルされました。音声ファイルの長さを参照せずに変換する場合は--noaudioオプションをつけて実行してください（非推奨）")
		logger.info('音声ファイル選択キャンセルにより終了')
		sys.exit(0)
	print(f'音声ファイル　　： {os.path.basename(wavpath)}')
	au = wave.open(wavpath, mode='rb')
	entire_xmax = au.getnframes() / au.getframerate()
	logger.info('正常終了：get_xmax_from_audio')
	return entire_xmax



def make_textgrid (item, isNoAudio):
	print(f'エクセルファイル： {os.path.basename(item)}')
	logger.info(f'{os.path.basename(item)} の処理開始')
	try:
		wb = xlrd.open_workbook(item)
	except Exception as e:
		print(f'{os.path.basename(item)} を開くことに失敗しました。')
		logger.exception(f'{e}')

	if isNoAudio == False:
		entire_xmax = get_xmax_from_audio(item)
	else:
		#音声ファイルを利用しない場合は最後列のxmaxを流用する
		xmax_list = wb.sheets()[0].col(1)
		entire_xmax = 0
		for i in range(1,len(xmax_list)):
			if "number" in str(xmax_list[i]):
				max_candidate = float(re.sub('number:', '', str(xmax_list[i])))
				entire_xmax = max_candidate

	#各種変数の初期化
	sheets = wb.sheets()
	sheet = sheets[0]

	#各種の処理をかけてtextgridのデータを作成する
	try:
		speaker_xmin_xmax_list = get_speaker_xmin_xmax_list(sheet)
		speaker_size = get_speaker_size(speaker_xmin_xmax_list)
		list_by_speaker = split_by_speaker(speaker_xmin_xmax_list)
		list_by_speaker = change_split_by_speaker_list(list_by_speaker, speaker_size)
		list_by_speaker = add_silent_section(list_by_speaker, speaker_size, entire_xmax)
	except Exception as e:
		logger.error(f'{os.path.basename(item)} の処理中にエラーが発生しました。')
		logger.exception(f'{e}')
		exit(1)
	try:
		data = compose_textgrid(entire_xmax, list_by_speaker, speaker_size)
	except Exception as e:
		logger.error(f'{os.path.basename(item)} の処理中にエラーが発生しました。')
		exit(1)

	#utf-8 で保存
	output_file_path = re.sub('.xls$|.xlsx$|.xlsm$', ".TextGrid", item)
	try:
		open(output_file_path, mode='w', encoding='utf_8').write(data)
		print(f'{os.path.basename(output_file_path)} が正常に出力されました')
		logger.info(f'{os.path.basename(item)} を {os.path.basename(output_file_path)} として保存に成功。')
	except Exception as e:
		print(f'{output_file_path} の保存に失敗しました。')
		logger.exception(f'{e}')



if __name__ == '__main__':
	logger.info('起動')
	args = argParse()
	print("エクセルファイル、続いて音声ファイル(wav)を選択してください。ファイル名が同一の場合は、音声ファイルは自動的に選択されます。同名のTextGridファイルが作成され、既存のファイルは上書きされますのでご注意ください。--serialオプションをつけると連続で処理できます。")
	logger.info('args: all - {}, serial - {}, noaudio - {}'.format(
		'true' if args.all else 'false',
		'true' if args.serial else 'false',
		'true' if args.noaudio else 'false',
	))

	# 処理するエクセルを取得する
	# --all かそれ以外かで場合分け
	if args.all == True:
		logger.info('「all」プロセス開始')
		collection = get_excel_collection()
		logger.info(f'処理対象：{", ".join(collection)}')
		for item in collection:
			# noaudio True
			make_textgrid(item, True)
		print("終了しました。")
		logger.info('「all」プロセス終了')
		exit(0)

	# --all でないならダイアログでひとつずつ選択
	else:
		item = get_excel_path()
		make_textgrid(item, args.noaudio)
		
		while args.serial == True:
			item = get_excel_path(item)
			make_textgrid(item, args.noaudio)

		print("終了しました。")
		logger.info('プロセス終了')
		exit(0)