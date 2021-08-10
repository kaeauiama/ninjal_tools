/**
 * 各種タグをチェックしてリストを表示する
 *
 */
 function lookUpTag() {
    // モーダルダイアログにテーブルを表示してＲタグの対応を決定する
    let html = HtmlService.createHtmlOutputFromFile("tagdialog");
    SpreadsheetApp.getUi().showModalDialog(html, "タグの中身リスト");
  }
  
  function getTagList() {
    // スプレッドシートからデータを取得
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = getData(sheet);
    if (data == null) return null; 
    let {standard_t} = data;

    if (!standard_t) {
      Browser.msgBox("データの取得に失敗しました。標準語テキストが必要です。", Browser.Buttons.OK);
      return null;
    }
  
    // 半角=>全角置換
    standard_t = standard_t.map(x=>x.toFullWidth());
    
    let lists = [
      {name: "Ｄタグ", list: makeTagList("（Ｄ：(.*?)）", standard_t)},
      {name: "Ｆタグ", list: makeTagList("（Ｆ：(.*?)）", standard_t)},
      {name: "Ｇタグ", list: makeTagList("（Ｇ：(.*?)）", standard_t)},
      {name: "Ｏタグ", list: makeTagList("（Ｏ：(.*?)）", standard_t)},
      {name: "Ｒタグ", list: makeTagList("（Ｒ：(.*?)）", standard_t)},
      {name: "Ｚタグ", list: makeTagList("（Ｚ：(.*?)）", standard_t)},
      {name: "人称タグ", list: makeTagList("（.(?:ＳＧ|ＰＬ|ＤＵ)：(.*?)）", standard_t)},
      {name: "山かっこ", list: makeTagList("＜(.*?)＞", standard_t)}
    ];
    
    console.log(lists);
    return lists;
  }
  
  const makeTagList = (key, data) => {
    // タグに入っている要素を重複なしで配列に入れる
    let list = [];
    const re = new RegExp(key, "g");
    for(let i=0,i_len=data.length; i<i_len; i++){
      let ar = re.exec(data[i]);
      while (ar) {
        list.push(ar[1]);
        ar = re.exec(data[i]);
      }
      re.lastIndex = 0; // lastIndex を初期化しないとバグる
    }
  
    list = list.filter(function (x,i,self){return self.indexOf(x) === i}); // 重複削除
    return list;
  }
  