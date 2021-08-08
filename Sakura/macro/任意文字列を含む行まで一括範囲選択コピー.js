var WshShell = new ActiveXObject("WScript.Shell");

var selStRow=0,selEdRow=0,curRow=0,hitRow=0,fileEndRow=0,searchStr="",rowSelect=false;

try {
  searchStr = InputBox("Input word","",64);
  //MessageBox(searchStr, 0);		// for DEBUG

  selStRow = GetSelectLineFrom;
  selEdRow = GetSelectLineTo;
  if (selStRow != selEdRow) { rowSelect = true; }

  if (!rowSelect) {
    GoLineTop;
    curRow = Number(ExpandParameter('$y'));
    
  } else {
    curRow = selStRow
    
  }
  
  SearchNext(searchStr, 4);
  hitRow = Number(ExpandParameter('$y'));

  GoFileEnd;
  fileEndRow = Number(ExpandParameter('$y'));
  
  Jump(curRow,0);
  SearchClearMark;

  BeginSelect;
  
  if (hitRow < fileEndRow-1) {
    Jump(hitRow+1,0);
  } else {
    GoFileEnd;
  }
  Copy;
} catch (error) {
  WshShell.Popup(error ,0,"Exception!",0);
  CancelMode;
  
}