/**
 * 方言テキストと標準語テキストの比較をする
 * 
 */
 function checkText() {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // ２列を取得
    let arrAlignData = sheet.getDataRange().getValues();
    let arrOriginal = sheet.getDataRange().getValues();
    const len = arrAlignData.length;
    
    // 半角=>全角置換
    // 人名X1みたいなのは過剰修正になるので要注意
    for(let i=0;i<len;i++){
      arrAlignData[i][0] = arrAlignData[i][0].toFullWidth();
      arrAlignData[i][1] = arrAlignData[i][1].toFullWidth();
    }
    
    // 端っこのスペース削除、ダブルスペースをシングルスペースに、タブを削除
    for(let i=0;i<len;i++){
      arrAlignData[i][0] = arrAlignData[i][0].replace(/^　*([^　].*[^　])　*$/g, "$1").replace(/　　/g, "　").replace(/\t/g, "");
      arrAlignData[i][1] = arrAlignData[i][1].replace(/^　*([^　].*[^　])　*$/g, "$1").replace(/　　/g, "　").replace(/\t/g, "");
    }
  
    // 整列オプションのかかり具合を差分チェックで確認
    for(let i=0;i<len;i++){
      let message = "";
      if(i==0){message+="半角全角・空白調整"; arrAlignData[i].push(message); continue}
      if(arrAlignData[i][0]=="" || arrAlignData[i][1]==""){message+="空セル"}
      if(arrAlignData[i][0]!=arrOriginal[i][0] || arrAlignData[i][1]!=arrOriginal[i][1]){message+="有"}
      arrAlignData[i].push(message);
    }
    
    // 記号の使い方があやしいもの
    for(let i=0;i<len;i++){
      let message = "";
      if(i==0){message+="記号系"; arrAlignData[i].push(message); continue}  
      if(arrAlignData[i][0].split('。').length!=arrAlignData[i][1].split('。').length){message+="句点個数 "}
      if(arrAlignData[i][0].split('　').length!=arrAlignData[i][1].split('　').length){message+="空白個数 "}
      if(arrAlignData[i][0].split('｛').length!=arrAlignData[i][1].split('｛').length){message+="｛｝個数 "}
      if(arrAlignData[i][1].split('｛').length!=arrAlignData[i][1].split('｝').length){message+="閉じミス "}
      if(arrAlignData[i][1].split('（').length!=arrAlignData[i][1].split('）').length){message+="閉じミス "}
      if(arrAlignData[i][1].split('＜').length!=arrAlignData[i][1].split('＞').length){message+="閉じミス "}
      if(/\n/.test(arrAlignData[i][0] + arrAlignData[i][1])){message+="改行有 "}
      if(/（[^：]+?）/.test(arrAlignData[i][1])){message+="：無括弧 "}
      if(/（[^（]+?　[^（]+?）/.test(arrAlignData[i][1])){message+="␣有カッコ "}
      if(/＜[^＝]+?＞/.test(arrAlignData[i][1])){message+="＝無山括弧 "}
      if(/（[^）]*?（(.*?)）(.*?)）/.test(arrAlignData[i][1])){message+="入れ子 "}
      if(/。[^　」]/.test(arrAlignData[i][0]+"　"+arrAlignData[i][1])){message+="句点後␣無 "}
      if(/[～＾”’！？↑↓、␣…＠【】]/.test(arrAlignData[i][0] + arrAlignData[i][1])){message+="ゴミ "}
      if(/[^ァ-ン0-9A-ZＡ-Ｚ。ー゜ｎ◆＊×　〔〕「」]/.test(arrAlignData[i][0].replace(/｛(.*?)｝/g, ""))){message+="方言ゴミ "}
      arrAlignData[i].push(message);
    }
    
    // タグ系
    for(let i=0;i<len;i++){
      let message = "";
      if(i==0){message+="タグ系"; arrAlignData[i].push(message); continue}  
      if(/(^|　)(私|わたし|あたし|俺|おれ|我|われ|あなた|貴方|あんた|君|きみ|お前|おまえ|儂|わし)/.test(arrAlignData[i][1].replace(/我慢/g, ""))){message+="人称 "}
      if(/(^|　)(おたく|お宅)/.test(arrAlignData[i][1])){message+="人称? "}
      if(/（３/.test(arrAlignData[i][1])){message+="３人称 "}
      if(/［[^］]{2,}］|［[^かがでとにのはへもをん]］/.test(arrAlignData[i][1].replace(/［から］/g, ""))){message+="謎省略 "}
      if(/（Ｄ：[ぁ-んー]+?）/.test(arrAlignData[i][1])){message+="Ｄかな "}
      if(/（Ｇ：[ぁ-んー]+?）/.test(arrAlignData[i][1])){message+="Ｇかな "}
      if(/（(?![ＤＦＧＫＯＲＺ]|[１２](ＳＧ|ＤＵ|ＰＬ))/.test(arrAlignData[i][1])){message+="謎タグ "}
      arrAlignData[i].push(message);
    }
      
    // Ｏタグ、Ｒタグの付け忘れを通知する
    let list = [];
    const re = new RegExp("（[ＯＲ]：(.*?)）", "g");
    for(let i=0; i<len; i++){
      let ar = re.exec(arrAlignData[i][1]);
      while (ar) {
        list.push(ar[1]);
        ar = re.exec(arrAlignData[i][1]);
      }
      re.lastIndex = 0; // lastIndex を初期化しないとバグる
    }
    list = list.filter(function (x,i,self){return self.indexOf(x) === i}); // 重複削除
    // いったん空白をpushして空白列をつくる
    for(let i=0;i<len;i++){
      if(i==0){arrAlignData[i].push("ＯＲタグ忘れ"); continue}
      arrAlignData[i].push("");
    }
    for(let i=0, listlen=list.length; i<listlen; i++){
      let reg = new RegExp("[^：]" + list[i]);
      for(let j=1;j<len;j++){
        let str = arrAlignData[j][1].replace(/（.+?）|＜|＝.+?＞/g, "");
        if(reg.test(str)){arrAlignData[j][5]+= list[i] + " "}      
      }
    }
    
    // 表記がおかしい可能性のあるもの
    for(let i=0;i<len;i++){
      let message = "";
      if(i==0){message+="表記"; arrAlignData[i].push(message); continue}
      if(/ぁ|ぃ|ぅ|ぇ|ぉ/.test(arrAlignData[i][1])){message+="小文字 "}
      if(/ーー/.test(arrAlignData[i][1])){message+="超長音 "}
      if(/[^カキクケコイィンタチツテト]゜/.test(arrAlignData[i][0])){message+="不正半濁点 "}
      // if(/＊/.test(arrAlignData[i][0] + arrAlignData[i][1])){message+="＊ "}
      if(/×/.test(arrAlignData[i][0] + arrAlignData[i][1])){message+="× "}
      if(/[アカガサザタダナハバパマヤラワ]ァ|[イキギシジチヂニヒビピミリ]ィ|[ウクグスズツヅヌフブプムユル]ゥ|[エケゲセセテデネヘベペメレ]ェ|[オコゴソゾトドノホボポモヨロヲ]ォ/.test(arrAlignData[i][0])){message+="カナ小文字 "}
      if(/[a-z]|[ａ-ｚ]/.test(arrAlignData[i][1])){message+="英小文字 "}
      if(/(^|[^XＸ0-9（])[0-9０-９]|[１２][^ＳＰＤ]/.test(arrAlignData[i][1])){message+="数字 "} // 人名X1のようなものと、人称１ＳＧのようなもの以外でアラビア数字を使ってはいけない。
      if(/ヅ/.test(arrAlignData[i][0])){message+="ヅ "}
      if(/ヂ/.test(arrAlignData[i][0])){message+="ヂ "}
      if(/[^　]ハ(　|$)/.test(arrAlignData[i][0])){message+="ハ "}
      if(/[^　]ヲ(　|$)/.test(arrAlignData[i][0])){message+="ヲ "}
      if(/[^　]ヘ(　|$)/.test(arrAlignData[i][0])){message+="ヘ "}
      if(/へ/.test(arrAlignData[i][0].replace(/｛(.*?)｝/g, ""))){message+="平仮名へ "}
      if(/べ/.test(arrAlignData[i][0].replace(/｛(.*?)｝/g, ""))){message+="平仮名べ "}
      if(/ｎ[ナニヌネノカキクケコ]/.test(arrAlignData[i][0])){message+="ｎ＋ナ行 "}
      arrAlignData[i].push(message);
    }
    
    // Ｋタグの存在
    for(let i=0;i<len;i++){
      let message = "";
      if(i==0){message+="Ｋタグ"; arrAlignData[i].push(message); continue}  
      if(/Ｋ：/.test(arrAlignData[i][1])){message+="Ｋ"}
      arrAlignData[i].push(message);
    }
    
    // 隠れた方言や口語的表現の正規化
    for(let i=0;i<len;i++){
      let message = "";
      let temp = arrAlignData[i][1];
        if(i==0){message+="口語・方言的"; arrAlignData[i].push(message); continue}
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
      arrAlignData[i].push(message);
    }
    
    // 文節に切るべきもの、つなげるべきものを検出
    for(let i=0;i<len;i++){
      let message = "";
      if(i==0){message+="文節の可能性"; arrAlignData[i].push(message); continue}
      // 方言テキスト
      if(/ナンカ/.test(arrAlignData[i][0])){message+="ナンカ "}
      // 標準語テキスト
      let temp = arrAlignData[i][1].replace(/［|］|（|）/g, '');
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
      arrAlignData[i].push(message);
    }
    
    // 精度の低いもの
    for(let i=0;i<len;i++){
      let message = "";
      if(i==0){message+="フィラー・する・なる"; arrAlignData[i].push(message); continue} 
      if(/(^|　)(ええ|えー|あの|その|この|こう|まあ|ま　|あー|んー|　ん　|うーん)/.test(arrAlignData[i][1])){message+="フィラー "}
      if(/[^　](する|して|した|しな|しま)/.test(arrAlignData[i][1].replace('そして', ''))){message+="する "}
      if(/[^　](なる|なった|なって|なら|なり|なれ|なろ)/.test(arrAlignData[i][1].replace(/[^　]*なら　/g, ''))){message+="なる "}
      arrAlignData[i].push(message);
    }
    
    // 引用形の総ざらい用  
    for(let i=0;i<len;i++){
      let message = "";
      let temp = arrAlignData[i][1].replace(/[　［］]/g, "");
      if(i==0){message+="引用"; arrAlignData[i].push(message); continue}
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
      if(/って　/.test(arrAlignData[i][1].replace('しまって', ''))){message+="って "}
      arrAlignData[i].push(message);
    }
    
    // オノマトペ用。現状はＧタグがついているかどうかにかかわらずすべて抽出する設定
    for(let i=0;i<len;i++){
      let message = "";
      if(i==0){message="オノマトペ類"; arrAlignData[i].push(message); continue}
      let OnomArr1 = arrAlignData[i][0].replace(/　/g, "").match(/([ぁ-んァ-ン]{2,5})\1/g) || [""];
      let OnomArr2 = arrAlignData[i][1].replace(/　/g, "").match(/([ぁ-んァ-ン]{2,5})\1/g) || [""];
      let OnomArr3 = arrAlignData[i][1].replace(/［］/g, "").match(/[^　]*(ーっ|ん)と/g) || [""];
      // 補足：matchメソッドではグローバルオプションgが設定されているときは、( )によるキャプチャは発生しない
      message += OnomArr1.reduce(function(x,y){return x + " " + y});
      message += OnomArr2.reduce(function(x,y){return x + " " + y});
      message += OnomArr3.reduce(function(x,y){return x + " " + y});
      arrAlignData[i].push(message);
    }
    
    // ゼロ準体方言用のチェックカラム
    for(let i=0;i<len;i++){
      let message = "";
      if(i==0){message+="準体言"; arrAlignData[i].push(message); continue}
      if(/[ただ][だで]/.test(arrAlignData[i][1])){message+="ただ "}
      if(/[のん][だで]/.test(arrAlignData[i][1])){message+="たのだ "}
      if(/[うくすつぬふむゆるぐずづぶ][だで]/.test(arrAlignData[i][1])){message+="るだ "}
      if(/[うくすつぬふむゆるぐずづぶ][のん][だで]/.test(arrAlignData[i][1])){message+="るのだ "}
      if(/[いー][だで]/.test(arrAlignData[i][1])){message+="いだ "}
      if(/[いー][のん][だで]/.test(arrAlignData[i][1])){message+="いのだ "}
      arrAlignData[i].push(message);
    }
  
    // チェック用
    for(let i=0;i<len;i++){
      let message = "";
      if(i==0){message+="タグ等"; arrAlignData[i].push(message); continue}
      if(/［/.test(arrAlignData[i][1])){message+="省 "}
      if(/（Ｄ/.test(arrAlignData[i][1])){message+="Ｄ "}
      if(/（Ｆ/.test(arrAlignData[i][1])){message+="Ｆ "}
      if(/（Ｇ/.test(arrAlignData[i][1])){message+="Ｇ "}
      if(/（Ｏ/.test(arrAlignData[i][1])){message+="Ｏ "}
      if(/（Ｒ/.test(arrAlignData[i][1])){message+="Ｒ "}
      if(/（Ｚ/.test(arrAlignData[i][1])){message+="Ｚ "}
      if(/（(１|２)/.test(arrAlignData[i][1])){message+="人 "}
      if(/＜/.test(arrAlignData[i][1])){message+="山 "}
      arrAlignData[i].push(message); // コメント：それぞれのタグで並べ替えできたら良いのだが……
    }
    
    // ひらがなが連続しすぎていたら注意
    for(let i=0;i<len;i++){
      let message = "";
      if(i==0){message+="仮名連"; arrAlignData[i].push(message); continue}
      let temp = arrAlignData[i][1].replace(/[　）［］＜]|｛(.*?)｝|（(.*?)：|＝(.*?)＞/g,"");
      let count = 0;
      let list = [];
      const re = /[^ぁ-ん]([ぁ-ん]+)[^ぁ-ん]/g;
      let ar = re.exec(temp);
      while (ar) {
        list.push(ar[1]);
        ar = re.exec(temp);
      }
      if (list == []) {
        arrAlignData[i].push("");
      }else{
        count = list.map(x=>x.length).reduce((acc,cur)=>Math.max(acc, cur), 0);
        message = count>=10 ? "平仮名"+count.toString()+"連続" : "";
        arrAlignData[i].push(message);
      }
      re.lastIndex = 0; // lastIndex を初期化しないとバグる
    }
    
    // 整理番号のフィル
    for(let i=0;i<len;i++){
      if(i==0){ arrAlignData[i].push("整理番号");
      }else{    arrAlignData[i].push(i);}
    }
    
    // すべて終わったらスプレッドシートに書き込む
    sheet.getRange(1,3,arrAlignData.length, arrAlignData[0].length).setValues(arrAlignData);
    return null;
  }