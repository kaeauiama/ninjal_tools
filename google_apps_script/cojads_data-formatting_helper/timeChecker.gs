/**
 * 発話区間の順序や妥当性について検討する
 * 
 */
 function timeChecker() {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    let {xmin_t, xmax_t, speaker_t, dialect_t, len, len_t} = getData(sheet);
  
    // 指摘事項をためる仕組み
    let checklists_t = [];
    let checklist = [];
    const processChecklist = () => {
      checklists_t.push(checklist);
      checklist = [];
    }
    
    // 発話区間長をとる
    let speechLengthList_t = [];
    for(let i=0;i<len;i++){
      let message = "";
      if(i==0){message+="発話区間長"; speechLengthList_t.push(message); continue}  
      speechLengthList_t.push(Math.round( (xmax_t[i] - xmin_t[i]) * 1000 ) / 1000);
    }
    
    // xmin<xmaxであるか／そもそもxmin, xmaxが存在するか確かめる
    for(let i=0;i<len;i++){
      let message = "";
      if(i==0){message+="発話区間"; checklist.push(message); continue}
      if(xmin_t[i] == null){message+="xminがない "}
      if(xmax_t[i] == null){message+="xmaxがない "}
      if(speechLengthList_t[i] < 0){message+="発話区間が負"}
      checklist.push(message);
    }
    processChecklist();
    
    // xminの昇順であるか確かめる[6]
    for(let i=0;i<len;i++){
      let message = "";
      if(i==0){message+="順序"; checklist.push(message); continue}  
      if(i+1 <= xmin_t.length){
        if(xmin_t[i] > xmin_t[i+1]){
          message+="下よりxminが遅い ";
        }
      }
      checklist.push(message);
    }
    processChecklist();
    
    // 話者ごとに発話区間が被っていないかチェックする
    let speakerList = [];
    checklist.push("各話者の被り");
    for(let i=1;i<len;i++){
      speakerList.push(speaker_t[i].toFullWidth());
      checklist.push("");
    }
    // 重複削除して話者リストをつくる
    speakerList = speakerList.filter(function (x,i,self){return self.indexOf(x) === i});
    
    for(let i=0;i<speakerList.length;i++){
      let speaker  = speakerList[i];
      let xmaxLast = 0;
      for(let j=1;j<len;j++){
        if(speaker_t[j].toFullWidth()==speaker){
          if (j==1) {
            xmaxLast = xmax_t[j];
            continue;
          }
          checklist[j] = xmaxLast == xmin_t[j]      ? "話者" + speaker + " 間隔ゼロ" 
                       : xmaxLast > xmin_t[j]       ? "話者" + speaker + " 区間重複"
                       : xmaxLast + 0.1 > xmin_t[j] ? "話者" + speaker + " 間隔が狭い(<0.1s)"
                                                    : "";
          xmaxLast = xmax_t[j];
        }
      }
    }
    processChecklist();
    
    // 発話区間が長すぎ・短すぎたら指摘する
    for(let i=0;i<len;i++){
      let message = "";
      if(i==0){message+="区間自体の長さ"; checklist.push(message); continue}
      if(speechLengthList_t[i] >= 10){message+="10s以上の区間"}
      if(speechLengthList_t[i] < 0.2){message+="0.2s未満の区間"}
      checklist.push(message);
    }
    processChecklist();
    
    // 話し手の発話速度平均を出す
    // （スペースと句点を抜いた文字数）／（xmax-xmin）で１秒当たり文字数として算出
    let totalTimespan  = 0;
    let totalCharacter = 0;
    for(let i=1;i<len;i++){
      totalTimespan  += speechLengthList_t[i];
      totalCharacter += dialect_t[i].replace(/　。「」｛｝/g, "").replace(/〔\d+?〕/g, "").length;
    }
    let avarageSpeechSpeedPerSecond = Math.round(totalCharacter * 1000 / totalTimespan) / 1000;
    Logger.log(totalTimespan + "   " + totalCharacter);
    
    
    // 発話区間長からの予測より大きく外れているものを検出する
    for(let i=0;i<len;i++){
      let message = "";
      if(i==0){message+="平均発話速度：\n" + avarageSpeechSpeedPerSecond + "字/s"; checklist.push(message); continue}
      let dialectLength = dialect_t[i].replace(/　。「」｛｝/g, "").replace(/〔\d+?〕/g, "").length;
      if(speechLengthList_t[i] * avarageSpeechSpeedPerSecond * 1.5 < dialectLength){message+="短"}
      if(speechLengthList_t[i] * avarageSpeechSpeedPerSecond * 2.0 < dialectLength){message+="すぎ"}
      if(dialectLength <= 5){
        if(speechLengthList_t[i] * avarageSpeechSpeedPerSecond * 0.3 > dialectLength){message+="長"}
        if(speechLengthList_t[i] * avarageSpeechSpeedPerSecond * 0.2 > dialectLength){message+="すぎ"}
      }else{
        if(speechLengthList_t[i] * avarageSpeechSpeedPerSecond * 0.5 > dialectLength){message+="長"}
        if(speechLengthList_t[i] * avarageSpeechSpeedPerSecond * 0.3 > dialectLength){message+="すぎ"}
      }
      checklist.push(message);
    }
    processChecklist();
    
    // 整理番号のフィル
    for(let i=0;i<len;i++){
      if(i==0){
        checklist.push("整理番号");
      }else{
        checklist.push(i);
      }
    }
    processChecklist();
    
    // すべて終わったらスプレッドシートに書き込む
    const speechLengthList = transpose([speechLengthList_t, ...checklists_t]);
    sheet.getRange(1, len_t+1, speechLengthList.length, speechLengthList[0].length).setValues(speechLengthList);
    return null;
  }