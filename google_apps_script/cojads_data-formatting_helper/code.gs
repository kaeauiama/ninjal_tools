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
   * 半角英数字を全角英数字に変換する
   * @name String#toFullWidth
   * @function
   * @param  {string} 任意の文字列
   * @return {string} 半角英数字部分が全角になった文字列
   */
  String.prototype.toFullWidth = function () {
    let str = this.replace(/[\!-\/\:-\`\[\]\{-\}]/g,
      function( tmpStr ) {
        return String.fromCharCode( tmpStr.charCodeAt(0) + 0xFEE0 );
      }
    );
    return str.replace(/ /g, "　");
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
  
  
  
  
  