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
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // 時間列を削ってテキスト列を取得
    let arrAlignData = sheet.getDataRange().getValues();
    arrAlignData.map(x=>x.splice(0,2));
  
    // 半角=>全角置換
    // 人名X1みたいなのは過剰修正になるので要注意
    for(let i in arrAlignData){
      arrAlignData[i][0] = arrAlignData[i][0].toFullWidth();
      arrAlignData[i][1] = arrAlignData[i][1].toFullWidth();
    }
    
    let lists = [
      {name: "Ｄタグ", list: makeTagList("（Ｄ：(.*?)）", arrAlignData)},
      {name: "Ｆタグ", list: makeTagList("（Ｆ：(.*?)）", arrAlignData)},
      {name: "Ｇタグ", list: makeTagList("（Ｇ：(.*?)）", arrAlignData)},
      {name: "Ｏタグ", list: makeTagList("（Ｏ：(.*?)）", arrAlignData)},
      {name: "Ｒタグ", list: makeTagList("（Ｒ：(.*?)）", arrAlignData)},
      {name: "Ｚタグ", list: makeTagList("（Ｚ：(.*?)）", arrAlignData)},
      {name: "人称タグ", list: makeTagList("（.(?:ＳＧ|ＰＬ|ＤＵ)：(.*?)）", arrAlignData)},
      {name: "山かっこ", list: makeTagList("＜(.*?)＞", arrAlignData)}
    ];
    
    console.log(lists);
    return lists;
  }
  
  const makeTagList = (key, data) => {
    // タグに入っている要素を重複なしで配列に入れる
    let list = [];
    const re = new RegExp(key, "g");
    for(let i=0,i_len=data.length; i<i_len; i++){
      let ar = re.exec(data[i][1]);
      while (ar) {
        list.push(ar[1]);
        ar = re.exec(data[i][1]);
      }
      re.lastIndex = 0; // lastIndex を初期化しないとバグる
    }
  
    list = list.filter(function (x,i,self){return self.indexOf(x) === i}); // 重複削除
    return list;
  }
  