/**
 * @file COJADS の規則に則ってテキスト修正案を出したり、必要なデータを抽出したりするスクリプト
 * @author Kazuki Aoyama <k.aoyama.macho@gmail.com>
 * @version v.2.0
 * @todo プロトタイプメソッドの JSDoc の書き方がよくわからない
 */

/**
 * 全角英数字を半角英数字に変換する
 * @name String#toHalfWidth
 * @function
 * @param  {string} 任意の文字列
 * @return {string} 全角英数字部分が半角になった文字列
 */
 String.prototype.toHalfWidth = function () {
    return this.replace(/[Ａ-Ｚａ-ｚ０-９]/g, function (char) {
      return String.fromCharCode(char.charCodeAt(0) - 65248);
    });
  }
  
  /**
   * 半角英数記号を全角に変換する
   * @name String#toFullWidth
   * @function
   * @param  {string} 任意の文字列
   * @return {string} 半角英数字部分が全角になった文字列
   */
  String.prototype.toFullWidth = function () {
    let str = this.replace(/[\!-\~]/g, char => String.fromCharCode( char.charCodeAt(0) + 0xFEE0 ));
    return str.replace(/\s|&nbsp;/g, "　");
  }
  
  // 転置行列をつくるワンライナー
  const transpose = (a) => a[0].map((_, c) => a.map((r) => r[c]));
  
  /**
   * ちゃんと入力ミスなく数値になっているかどうかチェックする
   * 数値 => true, 非数 => false
   */
  const isNum = (a) => {
    return isNaN(Number(a.toString())) ? false : true;
  }

  const FORMAT = {
    FOUR: 4,
    FIVE: 5,
    TIME: 9
  };

  /**
   * スプレッドシートのデータを取得する
   * xmin, xmax, 話者, 方言テキスト, 標準語テキストの５列を想定している
   * let {xmin_t, xmax_t, speaker_t, dialect_t, standard_t, len, len_t} = getData(); として使用できる
   */
  const getData = (sheet) => {
    const originalData = sheet.getDataRange().getValues();
    const len = originalData.length;
    const len_t = originalData[0].length;
    let original_t = transpose(originalData);

    // スプレッドシートのデータフォーマットを判定する
    const format = detectFormat(originalData[0]);
    if (format == null) {
      Browser.msgBox("フォーマットが判別できませんでした", Browser.Buttons.OK);
      return null;
    }

    // 検査対象のデータ
    if (format == FORMAT.FOUR) {
      return {
        format: FORMAT.FOUR,
        xmin_t: original_t[0],
        xmax_t: original_t[1],
        dialect_t: original_t[2],
        standard_t: original_t[3],
        len: len,
        len_t: len_t
      }
    }
    if (format == FORMAT.FIVE) {
      return {
        format: FORMAT.FIVE,
        xmin_t: original_t[0],
        xmax_t: original_t[1],
        speaker_t: original_t[2],
        dialect_t: original_t[3],
        standard_t: original_t[4],
        len: len,
        len_t: len_t
      }
    }
    if (format == FORMAT.TIME) {
      return {
        format: FORMAT.FIVE,
        xmin_t: original_t[0],
        xmax_t: original_t[1],
        speaker_t: original_t[2],
        dialect_t: original_t[3],
        len: len,
        len_t: len_t
      }
    }

  }

  /**
   * 貼り付けられたデータの構成を分析して
   * 旧プログラムとの互換性を担保する
   */
  const detectFormat = (header) => {
    const header4 = ["xmin","xmax","方言テキスト","標準語テキスト"];
    const header5 = ["xmin","xmax","話者","方言テキスト","標準語テキスト"];
    const header9 = ["xmin","xmax","話者","方言テキスト"];
    if (header == null || header.length <= 3) return null;
    if (header.length >= 5 && header.slice(0,4) == header5) return FORMAT.FIVE;
    if (header.slice(0,3) == header4) return FORMAT.FOUR;
    if (header.slice(0,3) == header9) return FORMAT.TIME;
    if (header.length == 4 && header[3].length <= 1) return FORMAT.FOUR;
    if (header.length == 4) return FORMAT.FOUR;
    if (header.length >= 5) return FORMAT.FIVE;
    return null;
  }
  
  /**
   * スプレッドシートにカスタムメニューを追加して UI からスクリプトを実行できるようにする
   */
  function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('カスタムメニュー')
      .addItem('現在表示されているシートに「修正個所チェック」', 'checkText')
      .addItem('現在表示されているシートに「Ｒタグ変換」', 'concealPrivacy')
      .addItem('現在表示されているシートに「タグ中身チェック」', 'lookUpTag')
      .addItem('現在表示されているシートに「自動補完検出」', 'autoCompleteDetector')
      .addItem('現在表示されているシートに「発話区間チェック」', 'timeChecker')
      .addItem('現在表示されているシートに「メタ情報付加」', 'metaMaker')
      .addItem('現在表示されているシートに「全体体裁チェック」', 'wholeChecker')
      .addToUi();
  }
  
  
  
  
  