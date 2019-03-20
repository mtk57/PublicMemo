// DeleteLineByKeyword.js
// キーワードにマッチする行を削除する
// 組み込みのSearchNext()がアホ(見つからない場合が分からない)なので、自力で検索しているww

// ★MUST
var KEYWORD = '\\t\\.git\\t';  // キーワード
var isReg = true;  // 正規表現ON/OFF


var shell = new ActiveXObject("WScript.Shell");

// 行頭に移動
GoFileTop();

// 行数を取得(0はおまじない)
var lineCount = GetLineCount(0);
//shell.Popup("lineCount=" + lineCount);

var line = 0;

//全行をループ
while (++line <= lineCount){
	var lineStr = GetLineStr(line);

	//shell.Popup("lineStr="+lineStr+"  line="+line+"  lineCount="+lineCount);

	if( isDeleteLine(lineStr, KEYWORD, isReg) ){
		DeleteLine();

		lineCount = GetLineCount(0);
		--line;

		//shell.Popup("lineStr="+lineStr+"  line="+line+"  lineCount="+lineCount);

		continue;
	}
	GoLineEnd();
	Right();
}
//shell.Popup("END");

function isDeleteLine(l, k, r) {
	if(r === true){
		var reg = new RegExp(k);
		return l.match(reg) != null;
	}
	else{
		return l.indexOf(k) != -1;
	}
}
