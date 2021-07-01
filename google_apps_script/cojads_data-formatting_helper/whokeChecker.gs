/**
 * 各種タグをチェックしてリストを表示する
 * @2020-07-30 Ｏ列に「方言チェック担当者」を追加することになったため、列数＋１など各所変更
 * @2021-01-08 要望に応じて「修正個所チェック」の機能をさらに移植してさらなるバグ減少を目指す
 *
 */
 function wholeChecker() {
    // サイドバーに修正点を表示する
    let html = HtmlService.createHtmlOutputFromFile("wholesidebar");
    SpreadsheetApp.getUi().showSidebar(html);
  }
  
  function getCorrectionList() {
    // データ取得
    const data = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getDataRange().getValues();
    const len = data.length;
    const data_t = transpose(data); // 転置行列
    let correctionList = [];
  
    try {
      // 行数チェック
      if (Math.max(data.map((x) => x.length)) > 30) {
        correctionList.push("右端に余計なデータがあるかも");
      }
  
      // タイトル行チェック
      const targetHeader = data[0];
      const correctHeader = [
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
  
      const excessItem = Array.from(dif(new Set(targetHeader), new Set(correctHeader)));
      const deficitItem = Array.from(dif(new Set(correctHeader), new Set(targetHeader)));
  
      if (targetHeader.length !== correctHeader.length) {
        correctionList.push("見出し行の長さが違う");
      }
      if (excessItem.length !== 0) {
        correctionList.push("見出し行に不要な要素「" + excessItem.join(",") + "」がある");
      }
      if (deficitItem.length !== 0) {
        correctionList.push("見出し行に必要な要素「" + deficitItem.join(",") + "」がない");
      }
      if (Array.from(dif(new Set(targetHeader), new Set(correctHeader))).length === 0 && targetHeader.join() !== correctHeader.join()) {
        correctionList.push("見出し行の要素の順序が違うかも");
      }
  
      // 空欄チェック
      for (let i = 1; i < len; i++) {
        for (let j = 0; j < 29; j++) {
          // 編集担当者、方言チェック担当者、注釈は空欄でよい
          if (j == 13 || j == 14 || (21 <= j && j <= 29)) continue;
          if (data[i][j] === "") {
            correctionList.push(`${(i + 1).toString()}行「${headerEnum[j]}」列が空`);
          }
        }
      }
  
      // メタ情報同一性チェック
      // 転置→重複排除して要素が２つ以上なら検出
      // xmin, xmax, ＩＤ, 話者, 方言, 標準語, 話者情報, 注釈以外は同一であるべき
      for (let i = 0; i < 28; i++) {
        if (i == 0 || i == 1 || (6 <= i && i <= 9) || (15 <= i && i <= 17) || (21 <= i && i <= 29)) continue;
        let tempArr = data_t[i].map((x) => x.toString()); // Date型はtoStringしておかないとset化しても重複排除されないっぽい
        tempArr.shift(); // 見出し行を取り除く
        let temp = Array.from(new Set(tempArr));
        if (temp.length > 1) {
          correctionList.push(`「${headerEnum[i]}」列の値が一様でない（${temp.join()}）`);
        }
      }
  
      // 県名チェック
      const ken = data[1][2];
      const kenList = [
        "北海道",
        "青森",
        "岩手",
        "宮城",
        "秋田",
        "山形",
        "福島",
        "茨城",
        "栃木",
        "群馬",
        "埼玉",
        "千葉",
        "東京",
        "神奈川",
        "新潟",
        "富山",
        "石川",
        "福井",
        "山梨",
        "長野",
        "岐阜",
        "静岡",
        "愛知",
        "三重",
        "滋賀",
        "京都",
        "大阪",
        "兵庫",
        "奈良",
        "和歌山",
        "鳥取",
        "島根",
        "岡山",
        "広島",
        "山口",
        "徳島",
        "香川",
        "愛媛",
        "高知",
        "福岡",
        "佐賀",
        "長崎",
        "熊本",
        "大分",
        "宮崎",
        "鹿児島",
        "沖縄",
      ];
      const kenListWithKen = [
        "青森県",
        "岩手県",
        "宮城県",
        "秋田県",
        "山形県",
        "福島県",
        "茨城県",
        "栃木県",
        "群馬県",
        "埼玉県",
        "千葉県",
        "神奈川県",
        "新潟県",
        "富山県",
        "石川県",
        "福井県",
        "山梨県",
        "長野県",
        "岐阜県",
        "静岡県",
        "愛知県",
        "三重県",
        "滋賀県",
        "兵庫県",
        "奈良県",
        "和歌山県",
        "鳥取県",
        "島根県",
        "岡山県",
        "広島県",
        "山口県",
        "徳島県",
        "香川県",
        "愛媛県",
        "高知県",
        "福岡県",
        "佐賀県",
        "長崎県",
        "熊本県",
        "大分県",
        "宮崎県",
        "鹿児島県",
        "沖縄県",
      ];
  
      if (kenList.indexOf(ken) != -1) {
      } else if (kenListWithKen.indexOf(ken) != -1) {
        correctionList.push(`「県」列に${ken}とあるが「県」不要`);
      } else if (ken == "東京都") {
        correctionList.push("「県」列に「東京都」とあるが「都」不要");
      } else if (ken == "大阪府" || ken == "京都府") {
        correctionList.push(`「県」列に${ken}とあるが「府」不要`);
      } else {
        correctionList.push("「県」列の文字列が変");
      }
  
      // xmin,xmax の入力ミスチェック
      const xminArray = data_t[0];
      const xmaxArray = data_t[1];
      xminArray.shift();
      xmaxArray.shift();
      for (let i = 1, i_len = xminArray.length; i < i_len; i++) {
        if (!isNum(xminArray[i])) correctionList.push(`${(i + 2).toString()}行：xminが不正`);
        if (!isNum(xmaxArray[i])) correctionList.push(`${(i + 2).toString()}行：xmaxが不正`);
      }
  
      // ＩＤ連番チェック
      const idArray = data_t[6];
      idArray.shift();
      for (let i = 1, i_len = idArray.length; i < i_len; i++) {
        if (idArray[i] != i + 1) correctionList.push(`${(i + 2).toString()}行：IDが${(i + 1).toString()}のはずが${idArray[i]}`);
      }
  
      // 修正支援ツールの一部を再掲
      const dialectTxt = data_t[8];
      const standardTxt = data_t[9];
      dialectTxt.shift();
      standardTxt.shift();
      for (let i = 1, i_len = dialectTxt.length; i < i_len; i++) {
        if (/\n/.test(dialectTxt[i] + standardTxt[i])) {
          correctionList.push(`${(i + 2).toString()}行：テキスト内に改行有`);
        }
        if (/\t/.test(dialectTxt[i] + standardTxt[i])) {
          correctionList.push(`${(i + 2).toString()}行：テキスト内にタブ有`);
        }
        if (/　　/.test(dialectTxt[i] + standardTxt[i])) {
          correctionList.push(`${(i + 2).toString()}行：テキスト内にダブルスペース有`);
        }
        if (/(^　|　$)/.test(dialectTxt[i] + standardTxt[i])) {
          correctionList.push(`${(i + 2).toString()}行：テキストの端にスペース有`);
        }
        if (dialectTxt[i] != dialectTxt[i].toFullWidth()) {
          correctionList.push(`${(i + 2).toString()}行：方言に半角文字有`);
        }
        if (standardTxt[i] != standardTxt[i].toFullWidth()) {
          correctionList.push(`${(i + 2).toString()}行：標準語に半角文字有`);
        }
        if (dialectTxt[i].split("。").length != standardTxt[i].split("。").length) {
          correctionList.push(`${(i + 2).toString()}行：句点の個数が異なる`);
        }
        if (dialectTxt[i].split("　").length != standardTxt[i].split("　").length) {
          correctionList.push(`${(i + 2).toString()}行：空白個数が異なる`);
        }
        if (dialectTxt[i].split("｛").length != standardTxt[i].split("｛").length) {
          correctionList.push(`${(i + 2).toString()}行：方言と標準語で波括弧の個数に齟齬がある`);
        }
        if (standardTxt[i].split("｛").length != standardTxt[i].split("｝").length) {
          correctionList.push(`${(i + 2).toString()}行：波括弧の閉じミス`);
        }
        if (standardTxt[i].split("（").length != standardTxt[i].split("）").length) {
          correctionList.push(`${(i + 2).toString()}行：丸括弧の閉じミス`);
        }
        if (standardTxt[i].split("＜").length != standardTxt[i].split("＞").length) {
          correctionList.push(`${(i + 2).toString()}行：山括弧の閉じミス`);
        }
        if (/（[^：]+?）/.test(standardTxt[i])) {
          correctionList.push(`${(i + 2).toString()}行：変な丸括弧有`);
        }
        if (/（[^（]+?　[^（]+?）/.test(standardTxt[i])) {
          correctionList.push(`${(i + 2).toString()}行：変な丸括弧有`);
        }
        if (/＜[^＝]+?＞/.test(standardTxt[i])) {
          correctionList.push(`${(i + 2).toString()}行：変な山括弧有`);
        }
        if (/（[^）]*?（(.*?)）(.*?)）/.test(standardTxt[i])) {
          correctionList.push(`${(i + 2).toString()}行：括弧の入れ子有`);
        }
        if (/。[^　」]/.test(dialectTxt[i] + "　" + standardTxt[i])) {
          correctionList.push(`${(i + 2).toString()}行：句点後に空白ナシ`);
        }
        if (/[～＾”’！？↑↓、␣…＠【】]/.test(dialectTxt[i] + standardTxt[i])) {
          correctionList.push(`${(i + 2).toString()}行：不要な記号有`);
        }
        if (/[^ァ-ン0-9A-ZＡ-Ｚ。ー゜ｎ◆＊×　〔〕「」]/.test(dialectTxt[i].replace(/｛(.*?)｝/g, ""))) {
          correctionList.push(`${(i + 2).toString()}行：方言に記号有`);
        }
      }
  
      // 話者とメタ情報の整合性チェック
      // 話者リストの作成
      let speakerArray = data_t[7].map((x) => x.toHalfWidth());
      speakerArray.shift();
      const speakerSet = new Set(speakerArray);
      const speakerList = Array.from(speakerSet);
      // 話者を key, メタ情報を value として連想配列化
      let speakerListAndMeta = {};
      loop: for (let item of speakerList) {
        for (let i = 0; i < len; i++) {
          if (data[i][7].toHalfWidth() === item) {
            speakerListAndMeta[item] = [data[i][15], data[i][16], data[i][17]];
            continue loop;
          }
        }
      }
      // チェックしていく
      for (let i = 1; i < len; i++) {
        if (speakerListAndMeta[data[i][7].toHalfWidth()].join() != [data[i][15], data[i][16], data[i][17]].join()) {
          correctionList.push(`${(i + 1).toString()}行：話者とメタ情報が不一致`);
        }
      }
  
      // 検出された情報を配列で返す
      let temptxt = "";
      for (let [key, value] of Object.entries(speakerListAndMeta)) {
        temptxt += `「${key}: ${value.join()}」`;
      }
      Logger.log(speakerList.join());
      Logger.log(temptxt);
    } catch (e) {
      correctionList = [e];
    }
    return correctionList;
  }
  
  const dif = (setA, setB) => {
    let _difference = new Set(setA);
    for (let elem of setB) {
      _difference.delete(elem);
    }
    return _difference;
  };
  
  const headerEnum = {
    0: "xmin",
    1: "xmax",
    2: "県",
    3: "地点",
    4: "file番号",
    5: "地点ID",
    6: "ID",
    7: "話者",
    8: "方言テキスト",
    9: "標準語テキスト",
    10: "データ名",
    11: "収録年月日",
    12: "収録場所",
    13: "編集担当者",
    14: "方言チェック担当者",
    15: "話者生年",
    16: "話者年齢",
    17: "話者性別",
    18: "話題",
    19: "収録担当者",
    20: "談話ジャンル",
    21: "注1番号",
    22: "注1語形",
    23: "注1注釈",
    24: "注2番号",
    25: "注2語形",
    26: "注2注釈",
    27: "注3番号",
    28: "注3語形",
    29: "注3注釈",
  };
  