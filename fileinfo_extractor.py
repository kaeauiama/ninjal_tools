# version 1.2
# last-modified 2021-04-27
# ------------------------------------------------------------------------
# viewer 用の metadata_squeezer からアイデアを流用（コードはほぼ全面書き換え）
# 同じフォルダにあるエクセルファイルからメタデータを抽出して CSV を吐く
# 地点名、ファイル名、収録年、長さ、話者数（総数、男、女）
# ------------------------------------------------------------------------
# TODO: 話者数を正確に出せるようにする

import glob, os, openpyxl, pandas, json, logging

# ログ出力用の設定
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
f1 = logging.Formatter('line %(lineno)d: [%(levelname)s] %(message)s')
stdout = logging.StreamHandler()
stdout.setLevel(logging.WARNING)
stdout.setFormatter(f1)
logger.addHandler(stdout)
f2 = logging.Formatter('%(asctime)s [line: %(lineno)03d] [%(levelname)s] %(message)s')
fileout = logging.FileHandler(filename="_fileinfo_extractor.log")
fileout.setLevel(logging.DEBUG)
fileout.setFormatter(f2)
logger.addHandler(fileout)
logger.propagate = False


# 県番号を返す
def get_prefecture(df):
	prefecture = df.iloc[0].loc["file番号"]
	logger.info(f'県番号：{prefecture}')
	return str(prefecture)


# 収録地点を返す
def get_location(df):
	pref = df.iloc[0].loc["県"]
	loc = df.iloc[0].loc["地点"]
	logger.info(f'地点：{pref+loc}')
	return pref + loc


# 収録年を返す
def get_recyear(df):
	recdate = df.iloc[0].loc["収録年月日"]
	# 「43532」のようなシリアル型は変換して年を取得
	if type(recdate) is int or type(recdate) is float:
		recyear = exceltime2datetime(recdate).year
	# 「YYYY年MM月DD日」形式は「年」より前の部分を取得
	elif type(recdate) is str:
		recdate = recdate.replace(" ", "")
		recyear = recdate.split("年")[0]
	else:
		logger.error(f'収録年が取得できません：recdate == {str(recdate)}: {type(recdate)}')
		return ""
	logger.info(f'収録年：{recyear}')
	return str(recyear)


# シリアル値 (int) → datetime 変換
# FROM: https://teratail.com/questions/186070
def exceltime2datetime(et):
    if et < 60:
        days = pandas.to_timedelta(et - 1, unit='days')
    else:
        days = pandas.to_timedelta(et - 2, unit='days')
    return pandas.to_datetime('1900/1/1') + days


# 収録時間を返す
# 現状ではエクセル上の最大のxmaxを利用してmm:ss文字列で返す
def get_length(df):
	xmax_max = df.loc[:,"xmax"].max()
	xmax_fmt = ""
	try:
		full_sec = int(xmax_max)
		hour, min, sec = full_sec// 3600, full_sec// 60 % 60, full_sec% 60
		xmax_fmt = '{HH}:{mm:02}:{ss:02}'.format(HH=hour, mm=min, ss=sec)
	except Exception as e:
		logger.error('型エラー：xmaxの最大値が数値型ではありません。')
		logger.error(f'xmax_max == {xmax_max}: {type(xmax_max)}')
		logger.exception(f'{e}')
	logger.info(f'長さ：{xmax_fmt}')
	return xmax_fmt


# 総話者数、男性数、女性数を返す
# 話者記号に揺れがある場合、数が多く出ることがある
def get_speaker(df, filename):
	speaker_df = df.loc[:,["話者", "話者性別"]].drop_duplicates().dropna()
	sum = len(speaker_df)
	male = len(speaker_df[speaker_df["話者性別"] == "男"])
	female = len(speaker_df[speaker_df["話者性別"] == "女"])
	logger.info(f'話者数：{sum}, {male}, {female}')
	if sum != male + female:
		logger.warning(f"{filename} - 注意：男女の合計が総数と違います。")
		logger.info(f'{speaker_df}')
	return str(sum), str(male), str(female)


# CSV として保存する
def save_as_csv(list):
	csv = ''
	for row in list:
		csv += ','.join(row) + '\n'

	try:
		open('metadata.csv', mode='w', encoding='utf-8').write(csv)
	except Exception as e:
		logger.exception(f'{e}')
		print('metadata.csv への書き込み失敗')
		print('処理終了')
		exit(1)
	logger.info('metadata.csv への書き込み成功')
	return None


# メイン関数
def extract_metadata (path):
	# 開いてデータフレームにする
	df = pandas.read_excel(path, dtype="object", engine='openpyxl')

	# ひとつずつ抽出する
	try:
		prefecture = get_prefecture(df)
		location = get_location(df)
		filename = os.path.splitext(os.path.basename(path))[0]
		logger.info(f'ファイル名：{filename}')
		recyear = get_recyear(df)
		length = get_length(df)
		speaker_sum, speaker_male, speaker_female = get_speaker(df, filename)
		metadata_list = [prefecture, location, "", "", filename, recyear, "", length, speaker_sum, speaker_male, speaker_female, "", ""]
		return metadata_list
	except Exception as e:
		print(f'{path} のメタデータ抽出失敗')
		logger.exception(f'{e}')
		return None


if __name__ == "__main__":
	logger.info('==========処理開始==========')

	excel_collection = sorted(glob.glob('*.xls*'))
	strlog = ", ".join(excel_collection)
	print(f'処理対象：{strlog}')
	logger.info(f'処理対象：{strlog}')

	header = ['prefecture','name','coordinate_x','coordinate_y','filename','year','release','length','speakers','male','female','detail','other']
	metadata_list = [header]

	# メタデータ抽出して貯める
	# 辞書型などではなく単純に順序固定のリストとして：
	# ファイル名、地点名、収録年、長さ、話者総数、話者男性、話者女性
	for item in excel_collection:
		print(f'処理中：{os.path.basename(item)}')
		logger.info(f'=====処理開始：{os.path.basename(item)}')
		metadata = extract_metadata(item)
		metadata_list.append(metadata)
		logger.info(f'=====処理終了：{os.path.basename(item)}')

	# 溜まったメタデータを CSV 出力する
	save_as_csv (metadata_list)

	print('処理が終了しました')
	logger.info('==========処理終了==========')