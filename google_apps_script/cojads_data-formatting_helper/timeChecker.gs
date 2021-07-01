/**
 * 発話区間の順序や妥当性について検討する
 * 
 */
 function timeChecker() {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    // ４列（xmin=[0],xmax=[1],speaker=[2],dialect=[3]）取得する
    let originalData = sheet.getDataRange().getValues();
    const len = originalData.length;
  
    
    // 発話区間長をとる[4]
    for(let i=0;i<len;i++){
      let message = "";
      if(i==0){message+="発話区間長"; originalData[i].push(message); continue}  
      originalData[i].push(Math.round( (originalData[i][1] - originalData[i][0]) * 1000 ) / 1000);
    }
    
    
    // xmin<xmaxであるか／そもそもxmin, xmaxが存在するか確かめる[5]
    for(let i=0;i<len;i++){
      let message = "";
      if(i==0){message+="発話区間"; originalData[i].push(message); continue}
      if(originalData[i][0] == null){message+="xminがない "}
      if(originalData[i][1] == null){message+="xmaxがない "}
      if(originalData[i][4] < 0){message+="発話区間が負"}
      originalData[i].push(message);
    }
    
    
    // xminの昇順であるか確かめる[6]
    for(let i=0;i<len;i++){
      let message = "";
      if(i==0){message+="順序"; originalData[i].push(message); continue}  
      if(originalData[i+1]){
        if(originalData[i][0] > originalData[i+1][0]){
          message+="下よりxminが遅い ";
        }
      }
      originalData[i].push(message);
    }
    
    
    // 話者ごとに発話区間が被っていないかチェックする[7]
    let speakerList = [];
    originalData[0].push("各話者の被り");
    for(let i=1;i<len;i++){
      speakerList.push(originalData[i][2].toHalfWidth());
      originalData[i].push(""); // 空白をpushしてあらかじめ[i][7]をつくっておく
    }
    // 重複削除して話者リストをつくる
    speakerList = speakerList.filter(function (x,i,self){return self.indexOf(x) === i});
    
    for(let i=0;i<speakerList.length;i++){
      let speaker  = speakerList[i];
      let xmaxLast = 0;
      for(let j=1;j<len;j++){
        if(originalData[j][2].toHalfWidth()==speaker){
          if (j==1) {
            xmaxLast = originalData[j][1];
            continue;
          }
          originalData[j][7] = xmaxLast == originalData[j][0]      ? "話者" + speaker + " 間隔ゼロ" 
                             : xmaxLast > originalData[j][0]       ? "話者" + speaker + " 区間重複"
                             : xmaxLast + 0.1 > originalData[j][0] ? "話者" + speaker + " 間隔が狭い(<0.1s)"
                                                                   : "";
          xmaxLast = originalData[j][1];
        }
      }
    }
    
    
    // 発話区間が長すぎ・短すぎたら指摘する[8]
    for(let i=0;i<len;i++){
      let message = "";
      if(i==0){message+="区間自体の長さ"; originalData[i].push(message); continue}
      if(originalData[i][4] >= 10){message+="10s以上の区間"}
      if(originalData[i][4] < 0.2){message+="0.2s未満の区間"}
      originalData[i].push(message);
    }
    
    
    // 話し手の発話速度平均を出す
    // （スペースと句点を抜いた文字数）／（xmax-xmin）で１秒当たり文字数として算出
    let totalTimespan  = 0;
    let totalCharacter = 0;
    for(let i=1;i<len;i++){
      totalTimespan  += originalData[i][4];
      totalCharacter += originalData[i][3].replace(/　。「」｛｝/g, "").replace(/〔\d+?〕/g, "").length;
    }
    let avarageSpeechSpeedPerSecond = Math.round(totalCharacter * 1000 / totalTimespan) / 1000;
    Logger.log(totalTimespan + "   " + totalCharacter);
    
    
    // 発話区間長からの予測より大きく外れているものを検出する[9]
    for(let i=0;i<len;i++){
      let message = "";
      if(i==0){message+="平均発話速度：\n" + avarageSpeechSpeedPerSecond + "字/s"; originalData[i].push(message); continue}
      let dialectLength = originalData[i][3].replace(/　。「」｛｝/g, "").replace(/〔\d+?〕/g, "").length;
      if(originalData[i][4] * avarageSpeechSpeedPerSecond * 1.5 < dialectLength){message+="短"}
      if(originalData[i][4] * avarageSpeechSpeedPerSecond * 2.0 < dialectLength){message+="すぎ"}
      if(dialectLength <= 5){
        if(originalData[i][4] * avarageSpeechSpeedPerSecond * 0.3 > dialectLength){message+="長"}
        if(originalData[i][4] * avarageSpeechSpeedPerSecond * 0.2 > dialectLength){message+="すぎ"}
      }else{
        if(originalData[i][4] * avarageSpeechSpeedPerSecond * 0.5 > dialectLength){message+="長"}
        if(originalData[i][4] * avarageSpeechSpeedPerSecond * 0.3 > dialectLength){message+="すぎ"}
      }
      originalData[i].push(message);
    }
    
    
    // 整理番号のフィル[10]
    for(let i=0;i<len;i++){
      if(i==0){
        originalData[i].push("整理番号");
      }else{
        originalData[i].push(i);
      }
    }
    
    
    // 話者リストの整形
    let arrangedSpeakerList = [];
    arrangedSpeakerList.push(["話者リスト\n"]);
    for(let i=0;i<speakerList.length;i++) arrangedSpeakerList.push([speakerList[i]]);
    
    
    // 話者全角半角混じりのチェック
    let flag = 0;
    const speakerStr = speakerList.join("");
    for(let i=1;i<len;i++){
      flag += speakerStr.indexOf(originalData[i][2])==-1 ? 1 : 0;
    }
    arrangedSpeakerList[0][0] += flag===0 ? "（全半角混じり無）" : "（全半角混じり有）";
    Logger.log("\nspeakerStr:" + speakerStr + "\nflag:" + flag);
    
    
    // すべて終わったらスプレッドシートに書き込む
    sheet.getRange(1,1,originalData.length, originalData[0].length).setValues(originalData);
    sheet.getRange(1, 1+originalData[0].length, arrangedSpeakerList.length, 1).setValues(arrangedSpeakerList);
    return null;
  }