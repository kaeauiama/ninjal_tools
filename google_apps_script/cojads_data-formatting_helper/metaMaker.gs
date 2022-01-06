const metaMaker = (data) => {
    // show modal dialog
    let html = HtmlService.createHtmlOutputFromFile("metamakerdialog");
    SpreadsheetApp.getUi().showModalDialog(html, "メタ情報の入力");
  }
  
  const getSpeakerList = () => {
    // 話者リストを返す
    const data = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getDataRange().getValues();
    const len = data.length;
    
    // 話者は data[i][7]
    let speakerList = [];
    for (let i=1;i<len;i++) speakerList.push(data[i][7]);
    speakerList = Array.from(new Set(speakerList));
   
    return speakerList;
  }
  
  const insertMetadata = (data) => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const value = sheet.getDataRange().getValues();
    
    const len = value.length;
    
    // prepare header
    value[0] = [
        "xmin",
        "xmax",
        "県",
        "地点",
        "file番号",
        "地点ID",
        "ID",
        "話者",
        "方言テキスト",
        "標準語テキスト",
        "データ名",
        "収録年月日",
        "収録場所",
        "編集担当者",
        "方言チェック担当者",
        "話者生年",
        "話者年齢",
        "話者性別",
        "話題",
        "収録担当者",
        "談話ジャンル",
        "注1番号",
        "注1語形",
        "注1注釈",
        "注2番号",
        "注2語形",
        "注2注釈",
        "注3番号",
        "注3語形",
        "注3注釈",
    ];
    
    // fill with blank
    for (let i=1;i<len;i++){
      value[i][10] = "";
      value[i][11] = "";
      value[i][12] = "";
      value[i][13] = "";
      value[i][14] = "";
      value[i][15] = "";
      value[i][16] = "";
      value[i][17] = "";
      value[i][18] = "";
      value[i][19] = "";
      value[i][20] = "";
      value[i][21] = "";
      value[i][22] = "";
      value[i][23] = "";
      value[i][24] = "";
      value[i][25] = "";
      value[i][26] = "";
      value[i][27] = "";
      value[i][28] = "";
      value[i][29] = "";
    }
    
    // metadata
    for (let i=1;i<len;i++){
      value[i][10] = data["reference"];
      value[i][11] = data["recordDate"];
      value[i][12] = data["recordPlace"];
      value[i][13] = data["editor"];
      value[i][14] = data["dialectChecker"];
      value[i][18] = data["topic"];
      value[i][19] = data["recorder"];
      value[i][20] = data["genre"];
    }
    
    // speakerdata
    const speakerList = getSpeakerList();
    for (let i=1;i<len;i++){
      value[i][15] = data[value[i][7]]["speakerBirthyear"];
      value[i][16] = data[value[i][7]]["speakerAge"];
      value[i][17] = data[value[i][7]]["speakerSex"];
    }
  
    // replace sheet
    console.log(value);
    sheet.getRange(1,1,value.length,value[0].length).setValues(value);
    return null;
  }
  