# version 4.4
# last-modified 2021-05-08
# --------------------------------------------------------------
# 校正作業の終了した Excel からコメント列等を削除し、
# それを COJADS サイト上で配布する CSV に変換し、ついでに ZIP を作成します
# デフォルトでは SJIS と UTF8 の CSV を作成するようになっています
# --sjis オプションをつけると sjis の CSV のみを作成します
# --utf8 オプションをつけると utf8 の CSV のみを作成します
# --------------------------------------------------------------

# -*- coding: utf-8 -*-
import argparse, glob, openpyxl, pandas, sys, re, io, os, zipfile, shutil, logging, traceback

# ログ出力用の設定
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
f1 = logging.Formatter('line %(lineno)d: [%(levelname)s] %(message)s')
stdout = logging.StreamHandler()
stdout.setLevel(logging.ERROR)
stdout.setFormatter(f1)
logger.addHandler(stdout)
f2 = logging.Formatter('%(asctime)s [%(filename)s: %(lineno)3d] [%(levelname)s] %(message)s')
fileout = logging.FileHandler(filename="_excel_to_csv.log")
fileout.setLevel(logging.DEBUG)
fileout.setFormatter(f2)
logger.addHandler(fileout)
logger.propagate = False


def noticeIllegalChar(e):
    errorMsg = traceback.format_exception_only(type(e), e)[0]
    try:
        regex = re.compile("(?<=character ')(.*)(?=' in)")
        char = regex.search(errorMsg).group(0)
        # Windows cmd はデフォルトで表示文字コードが cp932 なので、無難に unicode エンコードのままデコードせず表示する
        return "Shift-JISで表示できない文字'" + char + "'が含まれています（左の文字列で検索して、実際の文字を確認してください）。別の文字に置き換えるか、削除して再度お試しください。"
    except Exception as e:
        logger.warn(e)
        return "Shift-JISへの変換に失敗しました。"


def fullpath(path):
    base = os.path.dirname(os.path.abspath(__file__))
    return base + os.sep + path


def make_folder(path):
    if os.path.isdir(path):
        print("フォルダ " + path + " が既に存在します。上書きします。中止する場合は n を入力：", end="")
        str = input()
        if (str == "n"):
            logger.info("exited app due to user's interception")
            exit(0)
    else:
        os.mkdir(path)
        logger.info("created " + path + " directory")


def format_excel (path):
    # 不要な列を削除する
    df = pandas.read_excel(path, dtype="object", engine='openpyxl')
    expected_header = ["xmin","xmax","県","地点","file番号","地点ID","ID","話者","方言テキスト","標準語テキスト","データ名","収録年月日","収録場所","編集担当者","方言チェック担当者","話者生年","話者年齢","話者性別","話題","収録担当者","談話ジャンル","注1番号","注1語形","注1注釈","注2番号","注2語形","注2注釈","注3番号","注3語形","注3注釈"]
    actual_header = df.columns.values.tolist()
    for x in actual_header:
        if x not in expected_header:
            df=df.drop(columns=x)

    logger.info("dropped columns from " + path)

    # 収録年月日を「YYYY年 MM月 DD日」というフォーマットの文字列に変換
    # フォーマットについては考慮の余地あり
    recdate = format_datetime(df.iloc[0].loc['収録年月日'])
    cn = df.xmin.count() - 1
    df.loc[:cn,'収録年月日'] = recdate

    # 列数が足りないファイルは excel_error に移動して以降の処理に進ませない
    if len(df.columns) != 30:
        shutil.move(path, "excel_error" + os.sep + path)
        print('列数が30列に足りません。エラーフォルダに移動：', path)
        logger.warning("moved " + path + " to excel_error directory")
        return True

    # 古いファイルは excel_old フォルダに移動
    # 新しいファイルは excel_new フォルダに保存
    shutil.move(path, "excel_old" + os.sep + path)
    logger.info("moved " + path + " to excel_old directory")
    df.to_excel("excel_new" + os.sep + path, index=False)
    logger.info("added " + path + " to excel_new directory")
    return False


def format_datetime(value):
    date = value
    # シリアル型は変換
    if type(value) is int or type(value) is float:
        dt = exceltime2datetime(value)
        date = f'{str(dt.year)}年{str(dt.month)}月{str(dt.day)}日'
    # フォーマット
    date = date.replace(" ", "").replace("年", "年 ").replace("月", "月 ")
    return date


# シリアル値 (int) → datetime 変換
# FROM: https://teratail.com/questions/186070
def exceltime2datetime(et):
    if et < 60:
        days = pandas.to_timedelta(et - 1, unit='days')
    else:
        days = pandas.to_timedelta(et - 2, unit='days')
    return pandas.to_datetime('1900/1/1') + days


def load_excelfile_to_array (path):
    new_path = "excel_new" + os.sep + path
    sheet = openpyxl.load_workbook(new_path).active
    arr = []
    for r in range(1, sheet.max_row + 1):
        line = []
        for c in range(1, sheet.max_column + 1):
            if sheet.cell(r,c).value is None:
                line.append('')
            elif c == 7:
                line.append(str(sheet.cell(r,c).value).replace('\n','').replace('.0',''))
            else:
                line.append(str(sheet.cell(r,c).value).replace('\n',''))
        arr.append(line)
    logger.info("converted " + path + " to array")
    return arr


def save_array_to_csvfile (dataarr, csv_name, isOnlySjis, isOnlyUtf8):
    csv = ''
    for raw in dataarr:
        ln = ','.join(map(lambda x: re.sub("[\n,]", "", str(x)), raw))
        if ln == ',,,,,,,,,,,,,,,,,,,,,,,,,,,,,':
            continue
        csv += ln + '\n'

    # SJIS で保存する。--utf8 指定のときは飛ばす
    # Excel 内部は utf-8 なので、sjis として保存できない文字が存在する
    # SJIS 系文字コードは他に shift_jis, shift_jisx0213, shift_jis_2004 があるが
    # ここでは基本的な cp932 のみとする
    if isOnlyUtf8 == False:
        csv_name_sjis = 'csv_sjis' + os.sep + csv_name
        try:
            with open(csv_name_sjis + '.csv', mode='w', encoding='cp932') as w:
                w.write(csv)
            logger.info("saved " + csv_name + ".csv to csv_sjis directory")
        except Exception as e:
            # ユーザに変換不可文字を通知する
            logger.error(noticeIllegalChar(e))

    # UTF8 で保存する。--sjis 指定のときは飛ばす
    if isOnlySjis == False:
        csv_name_utf8 = 'csv_utf8' + os.sep + csv_name
        with open(csv_name_utf8 + '_utf8.csv', mode='w', encoding='utf-8') as w:
            w.write(csv)
        logger.info("saved " + csv_name + "_utf8.csv to csv_utf8 directory")


def excel_to_csv (sjis, utf8):
    excel_collection = glob.glob('*.xlsx') + glob.glob('*.xls')

    # 処理を行なうファイルを表示
    print(excel_collection)
    logger.info("csv list: [" + ",".join(excel_collection) + "]")

    err_count = 0

    # 不要列削除 -> 読み込み -> CSV として保存
    for excel_path in excel_collection:
        # excel_abspath = os.path.abspath(excel_path)
        try:
            err = format_excel (excel_path)
            if err:
                continue
        except Exception as e:
            print('columnのトリム処理失敗：', excel_path)
            logger.exception(f'{e}')
            err_count += 1
            continue
        try:
            excel_arr = load_excelfile_to_array (excel_path)
        except Exception as e:
            print('arrへの変換失敗：', excel_path)
            logger.exception(f'{e}')
            err_count += 1
            continue
        try:
            csv_name = re.sub(r'\.xlsx|\.xls','',excel_path)
            save_array_to_csvfile (excel_arr, csv_name, sjis, utf8)
        except Exception as e:
            print('csvへの変換失敗：', excel_path)
            logger.exception(f'{e}')
            err_count += 1
            continue
        print('変換終了：', excel_path)
    
    return err_count


def make_zip (isOnlyUtf8, isOnlySjis):
    csv_sjis_collection = sorted(glob.glob('csv_sjis' + os.sep + '*'))
    csv_utf8_collection = sorted(glob.glob('csv_utf8' + os.sep + '*'))

    # --utf8 指定のときは飛ばす
    if isOnlyUtf8 == False:
        z = zipfile.ZipFile("allcsv.zip", "w")
        for i in csv_sjis_collection:
            z.write(i)
        z.close()
        print('SJIS圧縮終了：allcsv.zip')
        logger.info("created allcsv.zip")

    # --sjis 指定のときは飛ばす
    if isOnlySjis == False:
        z = zipfile.ZipFile("allcsv_utf8.zip", "w")
        for i in csv_utf8_collection:
            z.write(i)
        z.close()
        print('UTF8圧縮終了：allcsv_utf8.zip')
        logger.info("created allcsv_utf8.zip")


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
    
    # フォルダを作成
    try: 
        make_folder("excel_error")
        make_folder("excel_old")
        make_folder("excel_new")
        make_folder("csv_sjis")
        make_folder("csv_utf8")
    except Exception as e:
        print("フォルダの作成に失敗しました。")
        logger.exception(f'{e}')

    # メイン処理へ渡す
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
            except Exception as e:
                print("ZIP の作成に失敗しました。")
                logger.exception(f'{e}')
    else:
        print('変換失敗したデータがあるため、ZIP 作成は行ないません。')
        
    print('処理が終了しました')
