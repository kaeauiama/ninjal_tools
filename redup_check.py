# version 2.0
# last-modified 2021-09-12
# ------------------------------------------------------------------------
# 重複箇所があるかどうかを確率で判定する
# 同一フォルダ内のファイルを読み込みレポートを作成する
# 同一県・同一地点で総当たり比較する
# データ前処理と重複判定は試行錯誤できるよう疎結合にする
# ------------------------------------------------------------------------

import argparse, csv, difflib, glob, itertools, logging, openpyxl, os, pandas, re
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

FILETYPE = ''
PROCESSMODE = ''

class TYPE(Enum):
    EXCEL = 'excel'
    CSV = 'csv'

class MODE(Enum):
    SPEEDY = 'speedy'
    NORMAL = 'normal'
    CAREFUL = 'careful'

# 設定
class CONFIG():
    THRESHOULD = 0.5

def check_encoding(filepath):
    detector = UniversalDetector()
    with open(filepath, mode='rb') as f:
        for binary in f:
            detector.feed(binary)
            if detector.done:
                break
    detector.close()
    return detector.result['encoding']


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
# x, y: 比較対象
# step: 比較のために切り出す文字列の長さ
# slide: 文字列を切り出す際のずらし幅（step の倍率）
# func: 類似度を算出するアルゴリズム
def exec_in_step(x, y, step, slide, func):
    x_list, y_list = [], []
    x_len, y_len = len(x), len(y)

    # 指定した文字数ごとに分割
    idx = 0
    while idx + step < x_len:
        x_list.append(x[idx:idx+step])
        idx += step * slide
    idx = 0
    while idx + step < y_len:
        y_list.append(y[idx:idx+step])
        idx += step * slide
    
    # 切り出した部分ごとに比較し、類似度最大を取得
    max_ratio = 0
    for x in x_list:
        for y in y_list:
            ratio = func(x, y)
            max_ratio = ratio if ratio > max_ratio else max_ratio
    return max_ratio


# 類似度を返す
def get_score(data1, data2):
    # モードによってサンプル抽出の設定値を変える
    step, slide = 0, 0
    global PROCESSMODE
    if PROCESSMODE == MODE.SPEEDY.value:
        step, slide = 200, 3
    elif PROCESSMODE == MODE.NORMAL.value:
        step, slide = 100, 1
    elif PROCESSMODE == MODE.CAREFUL.value:
        step, slide = 100, 1/4
    else:
        logger.error('モードが不明です')
        exit(1)

    # 任意の類似度算出メソッドを与える（現在はゲシュタルト）
    clazz = StringSimilarity()
    return exec_in_step(data1, data2, step, slide, clazz.gestalt)


# テキストを標準化する
# 標準語ならタグを取る処理が必要だが、方言で十分精度が出るので今回はこれだけ
def normalize(string):
    return re.sub('[　。\n\u3000]', '', string)


# データ抽出
def abstract_data(file_path):
    file_data = ''
    try:
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

    except Exception:
        logger.warning('WARN: データ抽出に失敗 filename=' + file_path)
        return None


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

    # 都道府県と地域ごとに整理する
    file_list = [];
    for pref in range(1,48):
        for place in 'abcdefghijklmnopqrstuvwxyz':
            list = []
            for file in file_collection:
                # 県番号と地域符号が同一のもの
                if format(pref, '02') == file[0:2] and place == file[3]:
                    # 場面設定でないもの
                    if '1' != file[5]:
                        list.append(file)
            if list != []:
                file_list.append({'pref':pref, 'place': place, 'files': list})
    return file_list


def argParse():
    argparser = argparse.ArgumentParser()
    argparser.add_argument('-f', '--file', choices=[e.value for e in TYPE], default=TYPE.CSV.value)
    argparser.add_argument('-m', '--mode', choices=[e.value for e in MODE], default=MODE.NORMAL.value)
    return argparser.parse_args()


def main():
    logger.info('==========処理開始==========')
    args = argParse()
    global FILETYPE, PROCESSMODE
    FILETYPE, PROCESSMODE = args.file, args.mode
    logger.info('対象：' + FILETYPE + ' モード：' + PROCESSMODE)

    # ファイル一覧を取得
    logger.info('ファイル一覧を取得中')
    file_list = get_file_collection()

    # 地点ごとに処理する
    suspicious_pair = []
    for item in file_list:
        pref = item['pref']
        place = item['place']
        file_list = item['files']
        logger.info('県番号: ' + format(pref, '02') + ' 地点: ' + place)
        logger.debug(item)

        # 当該地点にひとつしかファイルが無い場合はスルー
        if len(file_list) == 1:
            logger.info('　ファイルがひとつのため比較を行ないません')

        # 各データを総当たりで比較する
        data_list = []
        for item in file_list:
            data_list.append({'name': item, 'data': abstract_data(item)})

        for pair in itertools.combinations(data_list, 2):
            data1, data2 = pair[0], pair[1]
            if data1['data'] == None or data2['data'] == None:
                continue

            # 類似度を測定して表示する
            similarity = get_score(data1['data'], data2['data'])
            logger.info('　類似度：' + f'{similarity:.2f}' + '  比較ファイル：' + data1['name'] + ', ' + data2['name'])
            if (similarity > CONFIG.THRESHOULD):
                suspicious_pair.append({'pref': pref, 'place': place, 'file1': data1['name'], 'file2': data2['name'], 'similarity': similarity})

    logger.info('==========サマリ==========')
    if suspicious_pair == []:
        logger.info('類似度の高いファイルはありませんでした')
    else:
        logger.info('類似度の高いファイル：')
        for item in suspicious_pair:
            logger.info('　類似度: ' + str(item['similarity']) + ' 比較ファイル:' + item['file1'] + ', ' + item['file2'])

    logger.info('==========処理終了==========')


if __name__ == "__main__":
    main()