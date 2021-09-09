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

  const array_equal = (a, b) => {
    if (!Array.isArray(a))    return false;
    if (!Array.isArray(b))    return false;
    if (a.length != b.length) return false;
    for (var i = 0, n = a.length; i < n; ++i) {
      if (a[i] !== b[i]) return false;
    }
    return true;
  }

  const FORMAT = {
    XMIN_XMAX_SPE_DIA_STA: 1,
    XMIN_XMAX_SPE_DIA: 2,
    XMIN_XMAX_DIA_STA: 3,
    SPE_DIA_STA: 4,
    DIA_STA: 5,
    ERROR: 9
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
    if (format == FORMAT.ERROR) {
      Browser.msgBox("フォーマットが判別できませんでした", Browser.Buttons.OK);
      return null;
    }
    let data = {format: format, len: len, len_t: len_t};
    if (format == FORMAT.XMIN_XMAX_SPE_DIA_STA) {
        data["xmin_t"] = original_t[0];
        data["xmax_t"] = original_t[1];
        data["speaker_t"] = original_t[2];
        data["dialect_t"] = original_t[3];
        data["standard_t"] = original_t[4];
    }
    if (format == FORMAT.XMIN_XMAX_SPE_DIA) {
        data["xmin_t"] = original_t[0];
        data["xmax_t"] = original_t[1];
        data["speaker_t"] = original_t[2];
        data["dialect_t"] = original_t[3];
    }
    if (format == FORMAT.XMIN_XMAX_DIA_STA) {
        data["xmin_t"] = original_t[0];
        data["xmax_t"] = original_t[1];
        data["dialect_t"] = original_t[2];
        data["standard_t"] = original_t[3];
    }
    if (format == FORMAT.SPE_DIA_STA) {
        data["speaker_t"] = original_t[0];
        data["dialect_t"] = original_t[1];
        data["standard_t"] = original_t[2];
    }
    if (format == FORMAT.DIA_STA) {
        data["dialect_t"] = original_t[0];
        data["standard_t"] = original_t[1];
    }
    return data;
  }

  /**
   * 貼り付けられたデータの構成を分析して
   * 旧プログラムとの互換性を担保する
   */
  const detectFormat = (header) => {
    const header1 = ["xmin","xmax","話者","方言テキスト","標準語テキスト"];
    const header2 = ["xmin","xmax","話者","方言テキスト"];
    const header3 = ["xmin","xmax","方言テキスト","標準語テキスト"];
    const header4 = ["話者","方言テキスト","標準語テキスト"];
    const header5 = ["方言テキスト","標準語テキスト"];
    if (header == null || header.length <= 1) return FORMAT.ERROR;
    if (header.length >= 5 && array_equal(header.slice(0,5), header1)) return FORMAT.XMIN_XMAX_SPE_DIA_STA;
    if (header.length == 5) return FORMAT.XMIN_XMAX_SPE_DIA_STA;
    if (header.length >= 4 && array_equal(header.slice(0,4), header2)) return FORMAT.XMIN_XMAX_SPE_DIA;
    if (header.length >= 4 && array_equal(header.slice(0,4), header3)) return FORMAT.XMIN_XMAX_DIA_STA;
    if (header.length == 4) return FORMAT.XMIN_XMAX_DIA_STA;
    if (header.length >= 3 && array_equal(header.slice(0,3), header4)) return FORMAT.SPE_DIA_STA;
    if (header.length == 3) return FORMAT.SPE_DIA_STA;
    if (header.length >= 2 && array_equal(header.slice(0,2), header5)) return FORMAT.DIA_STA;
    if (header.length == 2) return FORMAT.DIA_STA;
    return FORMAT.ERROR;
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
  
  
  
  
  