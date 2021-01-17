# version 3.0
# last-modified 2020-04-20
# ------------------------------------------------------------------------
# Excel を TextGrid へと変換する
# TextGrid にない情報として音声の総時間が必要なので音声ファイルも読み込む
# バグフィックスした場合は Laravel アプリのほうにも反映させること
# ------------------------------------------------------------------------
# TODO: 前任者の設計のためリファクタリングの余地あり
# TODO: main 処理が長いのでもっと分離する


# -*- coding: utf-8 -*-
import argparse, codecs, xlrd, sys, re, io, os, tkinter, tkinter.filedialog, wave


#話者リストをつくる関数
def SpeakersSplit(speaker):
	#全角半角を統一する
	speakers_candidate_list1 = 'ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ'
	speakers_candidate_list2 = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
	speakers_list = []
	isAlphabet = True
	speaker = re.sub(r'[\'\s]', '', speaker)
	for i in range(len(speaker)):
		if speaker[i] in speakers_candidate_list1:
			speakers_list.append(speaker[i])
		elif speaker[i] in speakers_candidate_list2:
			speakers_list.append(chr(ord(speaker[i]) - ord('A') + ord('Ａ')))
		else:
			isAlphabet = False
	if isAlphabet == False:
		speakers_list.append(speaker)
	return speakers_list


#話者数を調べる関数
def GetSpeakerSize(speaker_xmin_xmax_list):
	size = 0
	checked_speaker = []
	for speaker, xmin, xmax, text in speaker_xmin_xmax_list:
		if speaker in checked_speaker:
			continue
		size += 1
		checked_speaker.append(speaker)
	return size


#話者ごとに配列を分けて、それをsplit_by_speaker_listにまとめる関数
#split_by_speaker_list[i]に各話者の発話データ（xmin, xmax, text）が入る
def SplitBySpeaker(speaker_xmin_xmax_list):
	split_by_speaker_list = []
	now_speaker = speaker_xmin_xmax_list[0][0]
	speaker_content = []
	for i in range(len(speaker_xmin_xmax_list)):
		if speaker_xmin_xmax_list[i][0] == now_speaker:
			speaker_content.append(speaker_xmin_xmax_list[i])
		else:
			split_by_speaker_list.append(speaker_content)
			speaker_content = []
			speaker_content.append(speaker_xmin_xmax_list[i])
			now_speaker = speaker_xmin_xmax_list[i][0]
	split_by_speaker_list.append(speaker_content)
	return split_by_speaker_list


#TextGridを書き出す関数
def MakeTextgridFile(file_path, xmax, split_by_speaker_list, speaker_size):
	with open(file_path, 'w') as output_file:
		print('File type = "ooTextFile"', file=output_file)
		print('Object class = "TextGrid"', file=output_file)
		print('', file=output_file)
		print('xmin = 0', file=output_file)
		print(f'xmax = {xmax}', file=output_file)
		print('tiers? <exists>', file=output_file)
		print(f'size = {speaker_size}', file=output_file)
		print('item []:', file=output_file)
		for i in range(speaker_size):
			print(f'    item [{i + 1}]:', file=output_file)
			print(f'        class = "IntervalTier"', file=output_file)
			print(f'        name = "{split_by_speaker_list[i][0][0]}"', file=output_file)
			print(f'        xmin = 0', file=output_file)
			print(f'        xmax = {xmax}', file=output_file)
			print(f'        intervals: size = {len(split_by_speaker_list[i])}', file=output_file)
			for j in range(len(split_by_speaker_list[i])):
				print(f'        intervals [{j + 1}]:', file=output_file)
				print(f'            xmin = {split_by_speaker_list[i][j][1]}', file=output_file)
				print(f'            xmax = {split_by_speaker_list[i][j][2]}', file=output_file)
				print(f'            text = {split_by_speaker_list[i][j][3]}', file=output_file)


#話者ごとにソートした際に第n行のxmaxが第n+1行のxminを越えて（＝発話区間が被って）しまうことがあるが、
#その場合に第n行のxmaxを狭めて発話区間の被りをなかったことにする関数
def ChangeSplitBySpeakerList(split_by_speaker_list, speaker_size):
	for i in range(speaker_size):
		for j in range(len(split_by_speaker_list[i]) - 1):
			if split_by_speaker_list[i][j][2] > split_by_speaker_list[i][j + 1][1]:
				speaker = split_by_speaker_list[i][j][0]
				xmin = split_by_speaker_list[i][j][1]
				xmax = split_by_speaker_list[i][j + 1][1]
				text = split_by_speaker_list[i][j][3]
				split_by_speaker_list[i][j] = (speaker, xmin, xmax, text)
	return split_by_speaker_list


#TextGridに空白区間を追加する関数
def AddSilentSection(split_by_speaker_list, speaker_size, entire_xmax):
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
	return split_by_speaker_list2


def argParse():
	argparser = argparse.ArgumentParser()
	argparser.add_argument("-s", "--serial", nargs="?", const=True, default=False, help="処理終了後、次のファイル選択画面を開きます")
	argparser.add_argument("-n", "--noaudio", nargs="?", const=True, default=False, help="音声ファイルの長さを参照せずにTextGridを作ります（非推奨）")
	return argparser.parse_args()


if __name__ == '__main__':
	args = argParse()

	print("エクセルファイル、続いて音声ファイル(wav)を選択してください。ファイル名が同一の場合は、音声ファイルは自動的に選択されます。同名のTextGridファイルが作成され、既存のファイルは上書きされますのでご注意ください。--serialオプションをつけると連続で処理できます。")

	# ダイアログを開いてエクセルファイルを選択
	root = tkinter.Tk()
	root.withdraw()
	fTyp = [("エクセルファイル","*.xlsx;*.xls;*.xlsm")]
	iDir = os.path.abspath(os.path.dirname(__file__))
	s = tkinter.filedialog.askopenfilename(filetypes = fTyp,initialdir = iDir)

	while True:
		if s=="":
			print("キャンセルされました。")
			sys.exit(1)
		print(f'エクセルファイル： {os.path.basename(s)}')
		wb = xlrd.open_workbook(s)

		#出力ファイル（textgrid）を作る
		output_file_path = re.sub('.xls$|.xlsx$|.xlsm$', ".TextGrid", s)

		if args.noaudio == False:
			#音声ファイルからxmaxを算出
			#同名のwavファイルがあるなら自動的に開く
			wavpath = re.sub('.xls$|.xlsx$|.xlsm$', ".wav", s)
			if not os.path.exists(wavpath):
				root = tkinter.Tk()
				root.withdraw()
				fTyp = [("音声ファイル","*.wav")]
				iDir = os.path.abspath(os.path.dirname(s))
				wavpath = tkinter.filedialog.askopenfilename(filetypes = fTyp,initialdir = iDir)
			if wavpath=="":
				print("キャンセルされました。音声ファイルの長さを参照せずに変換する場合は--noaudioオプションをつけて実行してください（非推奨）")
				sys.exit(1)
			print(f'音声ファイル　　： {os.path.basename(wavpath)}')
			au = wave.open(wavpath, mode='rb')
			entire_xmax = au.getnframes() / au.getframerate()
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
		xmin_list = sheet.col(0)
		xmax_list = sheet.col(1)
		speaker_list = sheet.col(7)
		text_list = sheet.col(8)

		#エラー判定
		if not (len(xmin_list) == len(xmax_list) and len(xmax_list) == len(speaker_list)):
			print('xmin, xmax, 話者の行数が揃っていないので揃えてください')
			sys.exit(1)

		#エクセルファイルから話者、xmin、xmax、方言テキストを抜き出したリストをつくり、話者でソートする
		speaker_xmin_xmax_list = []
		for i in range(1, len(xmin_list)):
			if not 'empty' in str(speaker_list[i]):
				speakers = re.sub('text:', '', str(speaker_list[i])) # text:'A' のように	typeとともに入っているからトリミングする
				speakers_list = SpeakersSplit(speakers)
				xmin = re.sub('number:', '', str(xmin_list[i]))
				xmax = re.sub('number:', '', str(xmax_list[i]))
				text = re.sub('text:', '', str(text_list[i]))
				text = re.sub(r'〔\d+?〕', '', text)
				text = re.sub('\\\\u3000', '　', text) #全角スペースがエスケープされているので戻す
				text = re.sub('\'', '"', text)
				for j in range(len(speakers_list)):
					speaker_xmin_xmax_list.append((speakers_list[j], float(xmin), float(xmax), text))
		speaker_xmin_xmax_list.sort()

		#各種の処理をかけてtextgridを作成する
		speaker_size = GetSpeakerSize(speaker_xmin_xmax_list)
		split_by_speaker_list = SplitBySpeaker(speaker_xmin_xmax_list)
		split_by_speaker_list = ChangeSplitBySpeakerList(split_by_speaker_list, speaker_size)
		split_by_speaker_list = AddSilentSection(split_by_speaker_list, speaker_size, entire_xmax)
		MakeTextgridFile(output_file_path + '_s', entire_xmax, split_by_speaker_list, speaker_size)

		#utf-8 に変換
		sf = open(output_file_path + '_s', mode='r', encoding='shift_jis').read()
		uf = open(output_file_path, mode='w', encoding='utf_8').write(sf)
		os.remove(output_file_path + '_s')
		print('正常に出力されました')

		# 引数--serial を指定しているときはループする
		if args.serial == False:
			break
		root = tkinter.Tk()
		root.withdraw()
		fTyp = [("エクセルファイル","*.xlsx;*.xls;*.xlsm")]
		s = tkinter.filedialog.askopenfilename(filetypes = fTyp, initialdir = os.path.abspath(os.path.dirname(s)))
