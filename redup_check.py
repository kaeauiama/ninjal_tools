# version 1.0
# last-modified 2021-09-09
# ------------------------------------------------------------------------
# 重複箇所があるかどうかを確率で判定する
# 同一フォルダ内のファイルを読み込みレポートを作成する
# 同一の都道府県内で比較する
# ふるさとことば集成を主として、調査に重複がないか調べる
# データ前処理と重複判定は試行錯誤できるよう疎結合にする
# ------------------------------------------------------------------------

import argparse, csv, difflib, glob, logging, openpyxl, os, pandas, re
from chardet.universaldetector import UniversalDetector
from enum import Enum
#import Levenshtein <- needs C++ 14 to install and build...


# ログ出力用の設定
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
f1 = logging.Formatter('line %(lineno)d: [%(levelname)s] %(message)s')
stdout = logging.StreamHandler()
stdout.setLevel(logging.INFO)
stdout.setFormatter(f1)
logger.addHandler(stdout)
f2 = logging.Formatter('%(asctime)s [line: %(lineno)03d] [%(levelname)s] %(message)s')
fileout = logging.FileHandler(filename="_redup_check.log")
fileout.setLevel(logging.DEBUG)
fileout.setFormatter(f2)
logger.addHandler(fileout)
logger.propagate = False

# 設定
THRESHOULD = 0.7
FILETYPE = ''

class TYPE(Enum):
    EXCEL = 'excel'
    CSV = 'csv'

def check_encoding(filepath):
    detector = UniversalDetector()
    with open(filepath, mode='rb') as f:
        for binary in f:
            detector.feed(binary)
            if detector.done:
                break
    detector.close()
    return detector.result['encoding']


def proc(str):
    ln = len(str)
    res = []
    idx = 0
    while idx + 2 <= ln:
        res.append(str[idx:idx+2])
        idx += 2
    print(res)
    return res


# ベクトル類似度をもとにした類似度測定法
class VectorSimilarity:
    # コサイン類似度
    def cos(self, x, y):
        pass

    # ピアソンの相関係数
    def pearson(self, x, y):
        pass

    # 偏差パタン類似度
    def hensa(self, x, y):
        pass


# 集合類似度をもとにした類似度測定法
class SetSimilarity:
    # Jaccard
    # X, Y が完全一致のときに 1 となる
    def jaccard(self, x, y):
        x = frozenset(x)
        y = frozenset(y)
        return len(x & y) / float(len(x | y))

    # Dice
    # X, Y が完全一致のときに 1 となる
    def dice(self, x, y):
        x = frozenset(x)
        y = frozenset(y)
        return 2 * len(x & y) / float(sum(map(len, (x, y))))

    # Simpson係数
    # X, Y が完全一致または部分一致のときに 1 となる
    def simpson(self, x, y):
        x = frozenset(x)
        y = frozenset(y)
        return len(x & y) / float(min(map(len, (x, y))))


# 文字列類似度をもとにした類似度測定法
class StringSimilarity:
    # ゲシュタルトパターンマッチング
    def gestalt(self, x, y):
        return difflib.SequenceMatcher(None, x, y).ratio()
 
"""
    # レーベンシュタイン距離
    def levenshtein(self, x, y):
        lev_dist = Levenshtein.distance(x, y)
        divider = len(x) if len(x) > len(y) else len(y)
        lev_dist = 1 - lev_dist / divider
        return lev_dist
 
    # ジャロ・ウィンクラー距離
    def jaro_winkler(self, x, y):
        jaro_dist = Levenshtein.jaro_winkler(x, y)
        return jaro_dist  
"""


# ずらし適用
# main が sub の中に現れるかどうかチェックする場合
def exec_in_step(main, sub, step, func):
    main_list, sub_list = [], []
    main_len, sub_len = len(main), len(sub)

    # 指定した文字数ごとに分割
    # sub は適当に重複させて切り出す
    idx = 0
    while idx + step < main_len:
        main_list.append(main[idx:idx+step])
        idx += step
    idx = 0
    while idx + step < sub_len:
        sub_list.append(sub[idx:idx+step])
        idx += int(step / 5)
    
    # 切り出した部分ごとに比較し、類似度最大を取得
    max_ratio = 0
    for x in main_list:
        for y in sub_list:
            ratio = func(x, y)
            max_ratio = ratio if ratio > max_ratio else max_ratio
    return max_ratio


# 類似度を返す
# ここを書き換えることで判定メソッドを入れ替える
def get_score(text, target):
    clazz = StringSimilarity()
    return exec_in_step(text, target, 100, clazz.gestalt)


# テキストを標準化する
# 標準語ならタグを取る処理が必要だが、方言で十分精度が出るので今回はこれだけ
def normalize(string):
    return re.sub('[　。\n\u3000]', '', string)


# データ抽出
def abstract_data(file_path):
    file_data = ''

    # ファイル形式毎に処理を分けて文字列データを取得
    if FILETYPE == TYPE.EXCEL.value:
        df = pandas.read_excel(file_path, dtype="object", engine='openpyxl')
        list = df['方言テキスト'].to_list()
        file_data = ''.join(list)

    elif FILETYPE == TYPE.CSV.value:
        enc = check_encoding(file_path)
        with open (file_path, 'r', encoding=enc) as f:
            for row in csv.DictReader(f):
                file_data += row['方言テキスト']

    # 標準化
    file_data_normalized = normalize(file_data)
    return file_data_normalized


# ファイルリストの整理
def get_file_collection():
    # 処理対象のファイルリストを取得
    file_collection = []
    if FILETYPE == TYPE.EXCEL.value:
        file_collection = sorted(glob.glob('*.xls*'))
    elif FILETYPE == TYPE.CSV.value:
        file_collection = sorted(glob.glob('*.csv*'))
    else:
        print('ERROR: ファイル形式が不明です。')
        exit(1)

    # 都道府県ごとに整理する
    file_list = [];
    for idx in range(1,48):
        furusato = ''
        chousa = []
        for csv_file in file_collection:
            if format(idx, '02') == csv_file[0:2]:
                # ふるさとことば集成
                if '099' == csv_file[5:8]:
                    furusato = csv_file
                # 調査（場面設定以外）
                elif '1' != csv_file[5]:
                    chousa.append(csv_file)
        file_list.append({'pref':idx, 'furusato': furusato, 'chousa': chousa})
    return file_list


def argParse():
    argparser = argparse.ArgumentParser()
    argparser.add_argument('-t', '--type', choices=['csv', 'excel'], default='csv')
    return argparser.parse_args()


def main():
    logger.info('==========処理開始==========')
    args = argParse()
    global FILETYPE
    FILETYPE = args.type

    # ファイル一覧を取得
    logger.info('ファイル一覧を取得中')
    file_list = get_file_collection()

    # 都道府県ごとに処理する
    logger.info('比較中')
    for item in file_list:
        pref_num = item['pref']
        furusato_file = item['furusato']
        chousa_file_list = item['chousa']
        logger.info('県番号: ' + format(pref_num, '02'))
        logger.info(item)
        
        # ふるさとことば集成のファイルが無い場合は無視
        if furusato_file == '':
            logger.info('INFO: ふるさとことば集成のデータがありません')
            continue
        # 調査ファイルが無い場合は無視
        if chousa_file_list ==[]:
            logger.info('INFO: 各地方言収集緊急調査のデータがありません')
            continue

        # ふるさとことば集成のデータを取得して加工する
        furusato_data = abstract_data(furusato_file)
        if furusato_data is None:
            logger.warning('WARN: ふるさとことば集成のデータ抽出に失敗 filename=' + furusato_file)
            continue

        # それ以外のデータを取得してひとつずつ類似度を測定する
        suspicious_files = []
        for chousa_file in chousa_file_list:
            chousa_data = abstract_data(chousa_file)
            if chousa_data is None:
                logger.warning('WARN: 各地方言収集緊急調査のデータ抽出に失敗 filename=' + chousa_file)
                continue

            # 類似度を測定して表示する
            similarity = get_score(furusato_data, chousa_data)
            logger.info('INFO: 類似度 ' + str(similarity) + '  filename=' + chousa_file)
            if (similarity > THRESHOULD):
                suspicious_files.append({'pref': pref_num, 'file': chousa_file, 'similarity': similarity})

    print('処理が終了しました')
    logger.info('==========サマリ==========')
    if suspicious_files == []:
        logger.info('類似度の高いファイルはありませんでした')
    else:
        logger.info('類似度の高いファイル：')
        for item in suspicious_files:
            logger.info('県番号: ' + str(item['pref']) + ' 類似度: ' + str(item['similarity']) + 'ファイル名:' + item['file'])

    logger.info('==========処理終了==========')


if __name__ == "__main__":
    main()