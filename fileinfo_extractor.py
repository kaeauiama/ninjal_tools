# version 1.2
# last-modified 2021-04-27
# ------------------------------------------------------------------------
# viewer 用の metadata_squeezer からアイデアを流用（コードはほぼ全面書き換え）
# 同じフォルダにあるエクセルファイルからメタデータを抽出して CSV を吐く
# 地点名、ファイル名、収録年、長さ、話者数（総数、男、女）
# ------------------------------------------------------------------------
# TODO: 話者数を正確に出せるようにする

import csv, glob, os, openpyxl, pandas, logging

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


# 指定した列の値を返す（一本値）
def get_attribute(col_name, df):
	value = df.iloc[0].loc[col_name]
	logger.info(f'{col_name}：{value}')
	return str(value)


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


# 話題を返す
def get_topic(df):
	topic_df = df.loc[:,["話題"]].drop_duplicates().dropna()
	topic_list = topic_df.values.flatten().tolist()
	logger.info(topic_list)
	topics = "、".join(topic_list)
	logger.info(f'話題：{topics}')
	return topics


# 総話者数、男性数、女性数を返す
# 話者記号に揺れがある場合、数が多く出ることがある
def get_speaker_num(df):
	speaker_df = df.loc[:,["話者", "話者性別"]].drop_duplicates().dropna()
	sum = len(speaker_df)
	male = len(speaker_df[speaker_df["話者性別"] == "男"])
	female = len(speaker_df[speaker_df["話者性別"] == "女"])
	unknown = len(speaker_df[speaker_df["話者性別"] == "不明"])
	logger.info(f'話者数：{sum}, {male}, {female}')
	if sum != male + female + unknown:
		logger.warning(f" - 注意：合計が総数と違います。")
		logger.info(f'{speaker_df}')
	return str(sum), str(male), str(female), str(unknown)


# CSV として保存する
def save_as_csv(filename, list):
	try:
		with open(f'{filename}.csv', mode='w', newline="", encoding='utf-8') as f:
			writer = csv.writer(f)
			writer.writerows(list)
	except Exception as e:
		logger.exception(f'{e}')
		print(f'{filename}.csv への書き込み失敗')
		print('処理終了')
		exit(1)
	print(f'{filename}.csv への書き込み成功')
	logger.info(f'{filename}.csv への書き込み成功')
	return None


# ファイル情報抽出メイン関数
def extract_filedata (path):
	# 開いてデータフレームにする
	df = pandas.read_excel(path, dtype="object", engine='openpyxl')

	# ひとつずつ抽出する
	try:
		prefecture = get_attribute('file番号', df)
		location = get_location(df)
		file_name = os.path.splitext(os.path.basename(path))[0]
		recyear = get_recyear(df)
		length = get_length(df)
		topic = get_topic(df)
		genre = get_attribute('談話ジャンル', df)
		speaker_sum, speaker_male, speaker_female, speaker_unknown = get_speaker_num(df)
		metadata_list = [prefecture, location, file_name, recyear, "", length, topic, genre, speaker_sum, speaker_male, speaker_female, speaker_unknown, "", ""]
		return metadata_list
	except Exception as e:
		print(f'{path} のメタデータ抽出失敗')
		logger.exception(f'{e}')
		return None


# 話者情報抽出メイン関数
def extract_speakerdata (path):
	# データフレーム抽出
	df = pandas.read_excel(path, dtype="object", engine='openpyxl')

	# 話者名抽出
	try:
		file_name = os.path.splitext(os.path.basename(path))[0]
		speaker_df = df.loc[:,["話者", "話者年齢", "話者性別"]].drop_duplicates().dropna()
		speaker_df.insert(0, "file_name", file_name, allow_duplicates=True)
		speaker_list = speaker_df.values.tolist()
		return speaker_list
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

	filedata_header = ['prefecture','location','file_name','record_year','release','length','topic','genre','speaker_num','speaker_male_num','speaker_female_num','speaker_unknown_num','detail','other']
	filedata_list = [filedata_header]

	speakerdata_header = ['file_name','name','age','sex']
	speakerdata_list = [speakerdata_header]

	# メタデータ抽出して貯める
	# 辞書型などではなく単純に順序固定のリストとして：
	# ファイル名、地点名、収録年、長さ、話者総数、話者男性、話者女性
	for item in excel_collection:
		print(f'処理中：{os.path.basename(item)}')
		logger.info(f'=====ファイル情報抽出処理　開始：{os.path.basename(item)}')
		filedata = extract_filedata(item)
		filedata_list.append(filedata)
		logger.info(f'=====ファイル情報抽出処理　終了：{os.path.basename(item)}')
		logger.info(f'=====話者情報抽出処理　開始：{os.path.basename(item)}')
		speakerdata = extract_speakerdata(item)
		speakerdata_list.extend(speakerdata)
		#for i in speakerdata:
		#	speakerdata_list.append(i)
		logger.info(f'=====話者情報抽出処理　終了：{os.path.basename(item)}')

	# 溜まったメタデータを CSV 出力する
	save_as_csv ('metadata_fileinfo_配布用', filedata_list)
	save_as_csv ('metadata_speakerinfo_配布用', speakerdata_list)

	print('処理が終了しました')
	logger.info('==========処理終了==========')