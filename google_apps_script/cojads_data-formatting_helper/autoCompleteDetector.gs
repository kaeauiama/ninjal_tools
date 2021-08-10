/**
 * 標準語テキストがまったく同内容のところを抜き出す
 * 「うん」みたいなありがちな標準語で始まっていたらあやしい
 */
 function autoCompleteDetector() {
    let html = HtmlService.createHtmlOutputFromFile("autocompletedialog");
    SpreadsheetApp.getUi().showModalDialog(html, "同一内容のテキストリスト");
  }
  
  const autoCompleteDetectorMain = () => {
    // 設定
    setting = {
      avoidShort: true,
      includeSimilar: false // あいまい検索はいまのところ使い物にならない
    };
  
    // データを取得
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = getData(sheet);
    if (data == null) return null; 
    let {standard_t, len} = data;
  
    if (!standard_t) {
      Browser.msgBox("データの取得に失敗しました。標準語テキストが必要です。", Browser.Buttons.OK);
      return null;
    }

    // 全角に置換しつつ標準語テキストを配列に入れる
    let text = standard_t.map(x=>x.toFullWidth());
  
    /**
     * 完全一致する２行を探す
     * たかだか O(1000) の比較なので愚直にやっても O(1000^2/2) = O(1000^2) で済む
     * 重複を避けるために一度検出したら referred に入れる
     * 結果は result_all に id list と text の連想配列で入れる
     */
    let referred = []; 
    let result_all = [];
    for(let i=0; i<len; i++) {
      if (referred.indexOf(i)!=-1) continue;
      let result_each = [];
      for(let j=i+1; j<len; j++) {
        if (setting.includeSimilar === false && text[i]===text[j]) {
          result_each.push(i, j);
          referred.push(i, j);
        }else if (setting.includeSimilar === true && vagueSearch(text[i],text[j]) === true) {
          result_each.push(i, j);
          referred.push(i, j);
        }
      }
      if (result_each.length>0) {
        if (setting.avoidShort == true && text[i].length < 7) continue;
        result_all.push({id: Array.from(new Set(result_each)), text: text[i]});
      }
    }
    if (result_all.length === []) {
      return result_all;
    } else {
      return false;
    }
  }
  
  
  /**
   * strA と strB に共通して含まれる文節を抽出する
   * その割合が長いほうの７割に達する場合は「あいまい一致」と判定する
   * 実用上、１文節の場合は即リターンするようにする
   */
  const vagueSearch = (strA, strB) => {
    const criterion = 0.7;
    const [long, short] = strA.length >= strB.length 
      ? [strA.split("　"), strB.split("　")] 
      : [strB.split("　"), strA.split("　")];
  
    
    if (long.length < 1) return false;
    
    let sharedBunsetsu = "";
    for (let i=0; i<short.length; i++) {
      if (long.indexOf(short[i])) {
        sharedBunsetsu += short[i];
      }
    }
  
    if (sharedBunsetsu.split("　").length / long.length > criterion) return true;
    else return false;
  }
  
  
  const test = () => {
    const a = "うん。　（Ｆ：あれ）　やっぱり　その　頃　馬耕［は］　あったのかな。";
    const b = "うん。　あれ　やっぱり　その　頃　馬耕［は］　あったのかな。";
    const result = vagueSearch(a, b);
    console.log(result);
  }