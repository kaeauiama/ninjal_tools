/**
 * 方言テキストと標準語テキストの比較をする
 * 
 */
 function checkText() {
    // データを取得
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = getData(sheet);
    if (data == null) return null; 
    let {speaker_t, dialect_t, standard_t, len, len_t} = data;

    if (!dialect_t || !standard_t) {
      Browser.msgBox("データの取得に失敗しました。方言テキストおよび標準語テキストが必要です。", Browser.Buttons.OK);
      return null;
    }

    let speaker_old = speaker_new = speaker_t;
    let dialect_old = dialect_new = dialect_t;
    let standard_old = standard_new = standard_t;

    // 指摘事項をためる仕組み
    let checklists_t = [];
    let checklist = [];
    const processChecklist = () => {
      checklists_t.push(checklist);
      checklist = [];
    }
    
    // 自動置換
    // 半角英数→全角に、端スペース削除、ダブルスペース→シングル、タブ・改行の削除
    const replaceIllegalPattern = (str) => {
      return str
        .toFullWidth()
        .replace(/^　*([^　].*[^　])　*$/g, "$1")
        .replace(/　　/g, "　")
        .replace(/[\t\n]/g, "");
    }
    dialect_new = dialect_new.map(x=>replaceIllegalPattern(x));
    standard_new = standard_new.map(x=>replaceIllegalPattern(x));

    // 話者記号を全角に
    if (speaker_t) {
      speaker_new = speaker_new.map(x => x.toFullWidth());
    }
  
    // 整列オプションのかかり具合を差分チェックで確認
    for(let i=0;i<len;i++){
      let message = "";
      if(i==0){message+="自動修正"; checklist.push(message); continue}
      if(dialect_new[i]=="" || standard_new[i]==""){message+="空セル "}
      if(dialect_new[i]!=dialect_old[i] || standard_new[i]!=standard_old[i]){message+="文修正有 "}
      if(speaker_t && speaker_new[i]!=speaker_old[i]){message+="話者修正有 "}
      checklist.push(message);
    }
    processChecklist();
    
    // 記号の使い方があやしいもの
    for(let i=0;i<len;i++){
      let message = "";
      if(i==0){message+="記号系"; checklist.push(message); continue}  
      if(dialect_new[i].split('。').length!=standard_new[i].split('。').length){message+="句点個数 "}
      if(dialect_new[i].split('　').length<standard_new[i].split('　').length){message+="標準語空白多 "}
      if(dialect_new[i].split('　').length>standard_new[i].split('　').length){message+="方言空白多 "}
      if(dialect_new[i].split('｛').length!=standard_new[i].split('｛').length){message+="｛｝個数 "}
      if(standard_new[i].split('｛').length!=standard_new[i].split('｝').length){message+="閉じミス "}
      if(standard_new[i].split('（').length!=standard_new[i].split('）').length){message+="閉じミス "}
      if(standard_new[i].split('＜').length!=standard_new[i].split('＞').length){message+="閉じミス "}
      if(/（[^：]+?）/.test(standard_new[i])){message+="：無括弧 "}
      if(/（[^（]+?　[^（]+?）/.test(standard_new[i])){message+="␣有カッコ "}
      if(/＜[^＝]+?＞/.test(standard_new[i])){message+="＝無山括弧 "}
      if(/（[^）]*?（(.*?)）(.*?)）/.test(standard_new[i])){message+="入れ子 "}
      if(/。[^　」]/.test(dialect_new[i]+"　"+standard_new[i])){message+="句点後␣無 "}
      if(/[～＾”’！？↑↓、␣…＠【】]/.test(dialect_new[i] + standard_new[i])){message+="ゴミ "}
      if(/[^ァ-ン0-9A-ZＡ-Ｚ。ー゜ｎ◆＊×　〔〕「」]/.test(dialect_new[i].replace(/｛(.*?)｝/g, ""))){message+="方言ゴミ "}
      checklist.push(message);
    }
    processChecklist();
    
    // タグ系
    for(let i=0;i<len;i++){
      let message = "";
      if(i==0){message+="タグ系"; checklist.push(message); continue}  
      if(/(^|　)(私|わたし|あたし|俺|おれ|我|われ|あなた|貴方|あんた|君|きみ|お前|おまえ|儂|わし)/.test(standard_new[i].replace(/我慢/g, ""))){message+="人称 "}
      if(/(^|　)(おたく|お宅)/.test(standard_new[i])){message+="人称? "}
      if(/（３/.test(standard_new[i])){message+="３人称 "}
      if(/［[^］]{2,}］|［[^かがでとにのはへもをん]］/.test(standard_new[i].replace(/［から］/g, ""))){message+="謎省略 "}
      if(/（Ｄ：[ぁ-んー]+?）/.test(standard_new[i])){message+="Ｄかな "}
      if(/（Ｇ：[ぁ-んー]+?）/.test(standard_new[i])){message+="Ｇかな "}
      if(/（(?![ＤＦＧＫＯＲＺ]|[１２](ＳＧ|ＤＵ|ＰＬ))/.test(standard_new[i])){message+="謎タグ "}
      checklist.push(message);
    }
    processChecklist();
      
    // Ｏタグ、Ｒタグの付け忘れを通知する
    let list = [];
    const re = new RegExp("（[ＯＲ]：(.*?)）", "g");
    for(let i=0; i<len; i++){
      let ar = re.exec(standard_new[i]);
      while (ar) {
        list.push(ar[1]);
        ar = re.exec(standard_new[i]);
      }
      re.lastIndex = 0; // lastIndex を初期化しないとバグる
    }
    list = list.filter(function (x,i,self){return self.indexOf(x) === i}); // 重複削除
    // いったん空白をpushして空白列をつくる
    for(let i=0;i<len;i++){
      if(i==0){checklist.push("ＯＲタグ忘れ"); continue}
      checklist.push("");
    }
    for(let i=0, listlen=list.length; i<listlen; i++){
      let reg = new RegExp("[^：]" + list[i]);
      for(let j=1;j<len;j++){
        let str = standard_new[j].replace(/（.+?）|＜|＝.+?＞/g, "");
        if(reg.test(str)){checklist[j]+= list[i] + " "}
      }
    }
    processChecklist();
    
    // 表記がおかしい可能性のあるもの
    for(let i=0;i<len;i++){
      let message = "";
      if(i==0){message+="表記"; checklist.push(message); continue}
      if(/ぁ|ぃ|ぅ|ぇ|ぉ/.test(standard_new[i])){message+="小文字 "}
      if(/ーー/.test(standard_new[i])){message+="超長音 "}
      if(/[^カキクケコイィンタチツテト]゜/.test(dialect_new[i])){message+="不正半濁点 "}
      // if(/＊/.test(dialect_new[i] + standard_new[i])){message+="＊ "}
      if(/×/.test(dialect_new[i] + standard_new[i])){message+="× "}
      if(/[アカガサザタダナハバパマヤラワ]ァ|[イキギシジチヂニヒビピミリ]ィ|[ウクグスズツヅヌフブプムユル]ゥ|[エケゲセセテデネヘベペメレ]ェ|[オコゴソゾトドノホボポモヨロヲ]ォ/.test(dialect_new[i])){message+="カナ小文字 "}
      if(/[a-z]|[ａ-ｚ]/.test(standard_new[i])){message+="英小文字 "}
      if(/(^|[^XＸ0-9（])[0-9０-９]|[１２][^ＳＰＤ]/.test(standard_new[i])){message+="数字 "} // 人名X1のようなものと、人称１ＳＧのようなもの以外でアラビア数字を使ってはいけない。
      if(/ヅ/.test(dialect_new[i])){message+="ヅ "}
      if(/ヂ/.test(dialect_new[i])){message+="ヂ "}
      if(/[^　]ハ(　|$)/.test(dialect_new[i])){message+="ハ "}
      if(/[^　]ヲ(　|$)/.test(dialect_new[i])){message+="ヲ "}
      if(/[^　]ヘ(　|$)/.test(dialect_new[i])){message+="ヘ "}
      if(/へ/.test(dialect_new[i].replace(/｛(.*?)｝/g, ""))){message+="平仮名へ "}
      if(/べ/.test(dialect_new[i].replace(/｛(.*?)｝/g, ""))){message+="平仮名べ "}
      if(/ｎ[ナニヌネノカキクケコ]/.test(dialect_new[i])){message+="ｎ＋ナ行 "}
      checklist.push(message);
    }
    processChecklist();
    
    // Ｋタグの存在
    for(let i=0;i<len;i++){
      let message = "";
      if(i==0){message+="Ｋタグ"; checklist.push(message); continue}  
      if(/Ｋ：/.test(standard_new[i])){message+="Ｋ"}
      checklist.push(message);
    }
    processChecklist();
    
    // 隠れた方言や口語的表現の正規化
    for(let i=0;i<len;i++){
      let message = "";
      let temp = standard_new[i];
        if(i==0){message+="口語・方言的"; checklist.push(message); continue}
      // 口語的
      if(/じゃ/.test(temp)){message+="じゃ "}
      if(/ちゃっ[たて]|ちゃ[いうえお]/.test(temp)){message+="ちゃう "}
      if(/[^建立た出で充当宛捨育持](てて|んでて|てない|んでない|いでない|てる|でる|てた|でた|てま|でま)/.test(temp)){message+="イ抜き "}
      if(/[いきぎじちびみり煮見えけげせぜてでねべめれ得出寝経]れ[たてなまる]/.test(temp)){message+="ラ抜き "}
      if(/りゃ/.test(temp)){message+="りゃ "}
      // 方言的
      if(/おらん|おる|おった|おって/.test(temp)){message+="おる "}
      if(/(^|　)よう　/.test(temp)){message+="よう "}
      if(/よる[よねわ。　]|よった[よねわ。　]/.test(temp)){message+="しよる "}
      if(/わ(　|。|$|な[^いかくけ]|ね|よ)/.test(temp)){message+="文末わ "}
      // if(/(の|のう)。/.test(temp)){message+="文末の?"}
      if(/[いきぎしじちみりっ](とり|とる|とった|とって)|ん(どり|どる|どった|どって)/.test(temp)){message+="しとる "}
      checklist.push(message);
    }
    processChecklist();
    
    // 文節に切るべきもの、つなげるべきものを検出
    for(let i=0;i<len;i++){
      let message = "";
      if(i==0){message+="文節の可能性"; checklist.push(message); continue}
      // 方言テキスト
      if(/ナンカ/.test(dialect_new[i])){message+="ナンカ "}
      // 標準語テキスト
      let temp = standard_new[i].replace(/［|］|（|）/g, '');
      if(/(^|　)(ああ|そう|こう|どう)([^　]|$)/.test(temp)){message+="そう "}
      if(/(^|　)(あ|そ|こ|ど)の([^　]|$)/.test(temp)){message+="その "}
      if(/(あ|そ|こ|ど)のような/.test(temp)){message+="そのような "}
      if(/かも　/.test(temp)){message+="かも "}
      if(/うち|宅|家/.test(temp)){message+="家 "}
      if(/[^　]　(よく|よう)　/.test(temp)){message+="よく "}
      if(/ければ　/.test(temp)){message+="なければ "}
      if(/[^　]ないと　/.test(temp)){message+="ないと "}
      if(/　なんて/.test(temp)){message+="なんて "}
      if(/(て|で)みれば/.test(temp)){message+="てみれば "}
      if(/(しょうが|しかた|しかたが|仕方|仕方が)　ない/.test(temp)){message+="仕方が "}
      if(/気に　入/.test(temp)){message+="気に入る "}
      if(/　(くらい|ぐらい)/.test(temp)){message+="くらい "}
      if(/[^　](わけ|ふう|もの([^　。]|$)|うちに)/.test(temp)){message+="形式名詞 "}
      if(/やっぱ/.test(temp)){message+="やっぱり "}
      if(/ばっか/.test(temp)){message+="ばっかり "}
      if(/みんな/.test(temp)){message+="みんな "}
      if(/あんま/.test(temp)){message+="あんまり "}
      if(/なんにも|何にも/.test(temp)){message+="なんにも "}
      if(/　だった/.test(temp)){message+="だった "}
      checklist.push(message);
    }
    processChecklist();
    
    // 精度の低いもの
    for(let i=0;i<len;i++){
      let message = "";
      if(i==0){message+="フィラー・する・なる"; checklist.push(message); continue} 
      if(/(^|　)(ええ|えー|あの|その|この|こう|まあ|ま　|あー|んー|　ん　|うーん)/.test(standard_new[i])){message+="フィラー "}
      if(/[^　](する|して|した|しな|しま)/.test(standard_new[i].replace('そして', ''))){message+="する "}
      if(/[^　](なる|なった|なって|なら|なり|なれ|なろ)/.test(standard_new[i].replace(/[^　]*なら　/g, ''))){message+="なる "}
      checklist.push(message);
    }
    processChecklist();
    
    // 引用形の総ざらい用  
    for(let i=0;i<len;i++){
      let message = "";
      let temp = standard_new[i].replace(/[　［］]/g, "");
      if(i==0){message+="引用"; checklist.push(message); continue}
      if(/という/.test(temp)){message+="という "}
      if(/と言う/.test(temp)){message+="と言う "}
      if(/といって/.test(temp)){message+="といって "}
      if(/と言って/.test(temp)){message+="と言って "}
      if(/と思う/.test(temp)){message+="と思う "}
      if(/と思って/.test(temp)){message+="と思って "}
      if(/そういう/.test(temp)){message+="そういう "}
      if(/そう言う/.test(temp)){message+="そう言う "}
      if(/そういって/.test(temp)){message+="そういって "}
      if(/そう言って/.test(temp)){message+="そう言って "}
      if(/そう思う/.test(temp)){message+="そう思う "}
      if(/そう思って/.test(temp)){message+="そう思って "}
      if(/って　/.test(standard_new[i].replace('しまって', ''))){message+="って "}
      checklist.push(message);
    }
    processChecklist();
    
    // オノマトペ用。現状はＧタグがついているかどうかにかかわらずすべて抽出する設定
    for(let i=0;i<len;i++){
      let message = "";
      if(i==0){message="オノマトペ類"; checklist.push(message); continue}
      let OnomArr1 = dialect_new[i].replace(/　/g, "").match(/([ぁ-んァ-ン]{2,5})\1/g) || [""];
      let OnomArr2 = standard_new[i].replace(/　/g, "").match(/([ぁ-んァ-ン]{2,5})\1/g) || [""];
      let OnomArr3 = standard_new[i].replace(/［］/g, "").match(/[^　]*(ーっ|ん)と/g) || [""];
      // 補足：matchメソッドではグローバルオプションgが設定されているときは、( )によるキャプチャは発生しない
      message += OnomArr1.reduce(function(x,y){return x + " " + y});
      message += OnomArr2.reduce(function(x,y){return x + " " + y});
      message += OnomArr3.reduce(function(x,y){return x + " " + y});
      checklist.push(message);
    }
    processChecklist();
    
    // ゼロ準体方言用のチェックカラム
    for(let i=0;i<len;i++){
      let message = "";
      if(i==0){message+="準体言"; checklist.push(message); continue}
      if(/[ただ][だで]/.test(standard_new[i])){message+="ただ "}
      if(/[のん][だで]/.test(standard_new[i])){message+="たのだ "}
      if(/[うくすつぬふむゆるぐずづぶ][だで]/.test(standard_new[i])){message+="るだ "}
      if(/[うくすつぬふむゆるぐずづぶ][のん][だで]/.test(standard_new[i])){message+="るのだ "}
      if(/[いー][だで]/.test(standard_new[i])){message+="いだ "}
      if(/[いー][のん][だで]/.test(standard_new[i])){message+="いのだ "}
      checklist.push(message);
    }
    processChecklist();
  
    // チェック用
    for(let i=0;i<len;i++){
      let message = "";
      if(i==0){message+="タグ等"; checklist.push(message); continue}
      if(/［/.test(standard_new[i])){message+="省 "}
      if(/（Ｄ/.test(standard_new[i])){message+="Ｄ "}
      if(/（Ｆ/.test(standard_new[i])){message+="Ｆ "}
      if(/（Ｇ/.test(standard_new[i])){message+="Ｇ "}
      if(/（Ｏ/.test(standard_new[i])){message+="Ｏ "}
      if(/（Ｒ/.test(standard_new[i])){message+="Ｒ "}
      if(/（Ｚ/.test(standard_new[i])){message+="Ｚ "}
      if(/（(１|２)/.test(standard_new[i])){message+="人 "}
      if(/＜/.test(standard_new[i])){message+="山 "}
      checklist.push(message); // コメント：それぞれのタグで並べ替えできたら良いのだが……
    }
    processChecklist();
    
    // ひらがなが連続しすぎていたら注意
    for(let i=0;i<len;i++){
      let message = "";
      if(i==0){message+="仮名連"; checklist.push(message); continue}
      let temp = standard_new[i].replace(/[　）［］＜]|｛(.*?)｝|（(.*?)：|＝(.*?)＞/g,"");
      let count = 0;
      let list = [];
      const re = /[^ぁ-ん]([ぁ-ん]+)[^ぁ-ん]/g;
      let ar = re.exec(temp);
      while (ar) {
        list.push(ar[1]);
        ar = re.exec(temp);
      }
      if (list == []) {
        checklist.push("");
      }else{
        count = list.map(x=>x.length).reduce((acc,cur)=>Math.max(acc, cur), 0);
        message = count>=10 ? "平仮名"+count.toString()+"連続" : "";
        checklist.push(message);
      }
      re.lastIndex = 0; // lastIndex を初期化しないとバグる
    }
    processChecklist();
    
    // 整理番号のフィル
    for(let i=0;i<len;i++){
      if(i==0){ checklist.push("整理番号");
      }else{    checklist.push(i);}
    }
    
    // すべて終わったらスプレッドシートに書き込む
    let modified;
    if (speaker_t){
      modified = transpose([speaker_new, dialect_new, standard_new, ...checklists_t]);
    } else {
      modified = transpose([dialect_new, standard_new, ...checklists_t]);
    }
    sheet.getRange(1, len_t+1, modified.length, modified[0].length).setValues(modified);
    return null;
  }