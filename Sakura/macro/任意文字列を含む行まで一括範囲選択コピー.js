// 任意文字列を含む行まで一括範囲選択コピー

var shell = new ActiveXObject("WScript.Shell");

var SEARCH_OPTION_REG = 4         // 検索オプション(正規表現)
var EX_PARAM_CURRENT_ROW = '$y'   // ExpandParameterの引数（現在の論理行位置(1始まり)を取得）
var MAX_SEARCH_WORD_LENGTH = 64


try{
  main();
}
catch (error) {
  shell.Popup(error ,0, "Exception!", 0);
  CancelMode;
}

function main(){
  var startRow = 0;
  var endRow = 0;
  var currentRow = 0;
  var hitRow = 0;
  var fileEndRow = 0;
  var searchWord = "";

  searchWord = InputBox("Input word", "", MAX_SEARCH_WORD_LENGTH);

  startRow = GetSelectLineFrom;
  endRow = GetSelectLineTo;

  if (startRow != endRow) {
    // シンプルにしたいので行選択中はNGとする
    MessageBox("It cannot be executed while a row is selected.", 0);
    return;
  }

  GoLineTop; 

  currentRow = Number(ExpandParameter(EX_PARAM_CURRENT_ROW));

  // 検索を実行
  // 「検索文字列」に指定した色で「検索マーク」が付く
  SearchNext(searchWord, SEARCH_OPTION_REG);

  hitRow = Number(ExpandParameter(EX_PARAM_CURRENT_ROW));
  
  GoFileEnd;

  fileEndRow = Number(ExpandParameter(EX_PARAM_CURRENT_ROW));
  
  Jump(currentRow, 0);

  SearchClearMark;
  
  BeginSelect;
  
  if (hitRow < fileEndRow - 1) {
    // 検索ヒット行がファイル最終行ではない場合
    Jump(hitRow + 1, 0);
  } else {
    // 検索文字列を含む行が[EOF]で終わってるか、[EOF]のみ行の一つ前の行に含む場合
    GoFileEnd;
  }
  
  Copy;
}