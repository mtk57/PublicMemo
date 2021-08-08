// 任意文字列を含む行まで一括範囲選択コピー
// ※用途に応じて検索条件に任意の文字列を変更してご利用ください。（デフォルトは改行のみ）
// ※指定行で折り返されてる場合、正常に動作しないので
// 　「設定(O)-折り返し方法(X)-折り返さない(X)」で実行してください。

var WshShell = new ActiveXObject("WScript.Shell");

var selStRow=0,selEdRow=0,curRow=0,hitRow=0,fileEndRow=0,searchStr="",rowSelect=false;

try {

  // 任意文字列を入力させる
  searchStr = InputBox("Input word","",64);

  //MessageBox(searchStr, 0);		// for DEBUG

  // 行選択チェック
  selStRow = GetSelectLineFrom; // 選択開始行取得
  selEdRow = GetSelectLineTo;   // 選択終了行取得
  if (selStRow != selEdRow) { rowSelect = true; }

  if (!rowSelect) {
    // 行を選択していない場合
    GoLineTop; // 行頭に移動(折り返し単位)
    curRow = Number(ExpandParameter('$y')); // 実行開始時の行
    
  } else {
    // 行を選択している場合
    curRow = selStRow
    
  }
  
  // 検索文字列を設定(お好みでどうぞ)
  //searchStr = "^(\r\n|\n)"; // 正規表現Crlfまたはlfのみの行
  //searchStr = "^[　 \t]*(\r\n|\n)"; // 全角・半角スペースやタブも空行とみなす場合
  //searchStr = "end if"; // 任意の文字列を含む行までとする場合(例:if～"end if")
  
  // 文字列検索　※正規表現で検索するため、検索オプションが変わるので注意。
  SearchNext(searchStr, 4);
  hitRow = Number(ExpandParameter('$y')); // 検索ヒット行取得
  
  // 最終行チェック
  GoFileEnd; // ファイルの最後に移動
  fileEndRow = Number(ExpandParameter('$y')); // ファイル最終行取得
  
  Jump(curRow,0);              // 現在行の先頭へジャンプ
  SearchClearMark;             // 検索マークの切替え(ハイライト解除)
  
  BeginSelect;                 // 範囲選択範囲選択モードオン
  
  if (hitRow < fileEndRow-1) { // 検索ヒット行がファイル最終行ではない場合
    Jump(hitRow+1,0);          // 検索ヒット行の一つ下先頭へジャンプ
    
    // 検索ヒットしない場合
    //if (preRow == hitRow) { GoFileEnd; } //ファイルの最後まで選択
    
  } else {
    // 検索文字列を含む行が[EOF]で終わってるか、[EOF]のみ行の一つ前の行に含む場合
    GoFileEnd; // ファイルの最後に移動
    
  }
  
  Copy; // 選択範囲をコピー
  
} catch (error) {
  WshShell.Popup(error ,0,"エラー",0);
  CancelMode; // 各種モードの取り消し
  
}