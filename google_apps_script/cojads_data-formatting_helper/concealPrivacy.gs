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
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // ２列をコピーしてから「整列オプション」をかけていく
    let arrAlignData = sheet.getDataRange().getValues();
  
    // 半角=>全角置換
    // 人名X1みたいなのは過剰修正になるので要注意
    for(let i in arrAlignData){
      arrAlignData[i][0] = arrAlignData[i][0].toFullWidth();
      arrAlignData[i][1] = arrAlignData[i][1].toFullWidth();
    }
    
    // Ｒタグに入っている要素を重複なしで配列に入れる
    let RTagList = [];
    const re = new RegExp("（Ｒ：(.*?)）", "g");
    for(let i=0,i_len=arrAlignData.length; i<i_len; i++){
      let ar = re.exec(arrAlignData[i][1]);
      while (ar) {
        RTagList.push(ar[1]);
        ar = re.exec(arrAlignData[i][1]);
      }
      re.lastIndex = 0; // lastIndex を初期化しないとバグる
    }
    if(RTagList.length == false){
      SpreadsheetApp.getActiveSpreadsheet().toast("Ｒタグはありませんでした。");
      return null;
    }
    RTagList = RTagList.filter(function (x,i,self){return self.indexOf(x) === i}); // 重複削除
  
    // 対応表を作る
    let RTagArr = [["対応表",""]];
    let temp = [];
    for(let i=1; i <= RTagList.length; i++){
      temp = [RTagList[i-1], "X" + i];
      RTagArr.push(temp);
    }
    Logger.log(RTagArr);
    return RTagArr;
  }
  
  function replaceRTag(RTagArr) {
    console.log(RTagArr);
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    let arrAlignData = sheet.getDataRange().getValues();
    for(let i in arrAlignData){
      arrAlignData[i][0] = arrAlignData[i][0].toFullWidth();
      arrAlignData[i][1] = arrAlignData[i][1].toFullWidth();
    }
    
    // 標準語のＲタグを書き換えて、Ｅ列に対応する文節を入れる
    for(let i=0,i_len=arrAlignData.length; i<i_len; i++){
      let str = "";
      if(i===0) {str = "Ｒタグ変換";}
      else{ 
        let [arrBunsetuDialect, arrBunsetuStandard] = arrAlignData[i][1].indexOf("　")===-1 
          ? [[arrAlignData[i][0]], [arrAlignData[i][1]]] 
          : [arrAlignData[i][0].split("　"), arrAlignData[i][1].split("　")];
        for(let j=1,j_len=RTagArr.length; j<j_len; j++){ // j=0 はタイトル行
          for(let k in arrBunsetuStandard){
            let fstr = "（Ｒ：" + RTagArr[j][0] + "）";
            while(arrBunsetuStandard[k].indexOf(fstr)!==-1){
              arrAlignData[i][1] = arrAlignData[i][1].replace(fstr, "（Ｒ：" + RTagArr[j][1] + "）");
              str += arrBunsetuDialect[k];
              arrBunsetuStandard[k] = "";
            }
          }
        }
      }
      arrAlignData[i].push(str);
    }
  
    // すべて終わったらスプレッドシートに書き込む 
    sheet.getRange(1,3,arrAlignData.length, 3).setValues(arrAlignData);
    sheet.getRange(1, 6, RTagArr.length, 2).setValues(RTagArr);
    return null;
  }