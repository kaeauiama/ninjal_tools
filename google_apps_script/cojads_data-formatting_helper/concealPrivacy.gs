/**
 * Ｒタグをマスクするための関数
 * 標準語テキストのＲタグ内を順番に X1, X2,... で置換し、次列にその目印をつける
 * 方言テキストの自動置換は無理なので、のちにその目印をめやすに手動で置換する
 */
 function concealPrivacy() {
    // モーダルダイアログにテーブルを表示してＲタグの対応を決定する
    let html = HtmlService.createHtmlOutputFromFile("rtagdialog");
    SpreadsheetApp.getUi().showModalDialog(html, "Ｒタグ変換対応");
  }
  
  function getRTagArr() {
    // スプレッドシートからデータを取得
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    let {standard_t} = getData(sheet);
  
    // 半角=>全角置換
    standard_t = standard_t.map(x=>x.toFullWidth());
    
    // Ｒタグに入っている要素を重複なしで配列に入れる
    let RTagList = [];
    const re = new RegExp("（Ｒ：(.*?)）", "g");
    standard_t.forEach(x=>{
      let ar = re.exec(x);
      while (ar) {
        RTagList.push(ar[1]);
        ar = re.exec(x);
      }
      re.lastIndex = 0; // lastIndex を初期化しないとバグる
    });
    if(RTagList.length == false){
      return false;
    }
    RTagList = RTagList.filter(function (x,i,self){return self.indexOf(x) === i}); // 重複削除
  
    // 対応表を作る
    let RTagArr = [["対応表",""]];
    let temp = [];
    for(let i=1; i <= RTagList.length; i++){
      temp = [RTagList[i-1], "Ｘ" + i.toString().toFullWidth()];
      RTagArr.push(temp);
    }
    Logger.log(RTagArr);
    return RTagArr;
  }
  
  function replaceRTag(RTagArr) {
    console.log(RTagArr);
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    let {dialect_t, standard_t, len, len_t} = getData(sheet);
    dialect_t = dialect_t.map(x=>x.toFullWidth());
    standard_t = standard_t.map(x=>x.toFullWidth());
    let checklist_t = [];
    
    // 標準語のＲタグを書き換えて、対応する文節を入れる
    for(let i=0; i<len; i++){
      let str = "";
      if(i===0) {str = "Ｒタグ変換";}
      else{ 
        let [arrBunsetuDialect, arrBunsetuStandard] = standard_t[i].indexOf("　")===-1 
          ? [[dialect_t[i]], [standard_t[i]]] 
          : [dialect_t[i].split("　"), standard_t[i].split("　")];
        for(let j=1,j_len=RTagArr.length; j<j_len; j++){ // j=0 はタイトル行
          for(let k in arrBunsetuStandard){
            let fstr = "（Ｒ：" + RTagArr[j][0] + "）";
            while(arrBunsetuStandard[k].indexOf(fstr)!==-1){
              standard_t[i] = standard_t[i].replace(fstr, "（Ｒ：" + RTagArr[j][1] + "）");
              str += arrBunsetuDialect[k];
              arrBunsetuStandard[k] = "";
            }
          }
        }
      }
      checklist_t.push(str);
    }

    // 書き込むデータを整形する
    const modified = transpose([dialect_t, standard_t, checklist_t]);
  
    // スプレッドシートに書き込む 
    sheet.getRange(1, len_t+1,modified.length, modified[0].length).setValues(modified);
    sheet.getRange(1, len_t+modified[0].length+1, RTagArr.length, RTagArr[0].length).setValues(RTagArr);
    return null;
  }