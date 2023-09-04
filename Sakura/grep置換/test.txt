set APP="%ProgramFiles(x86)%\sakura\sakura.exe"
set TARGET="C:\_git\PublicMemo\Sakura\grep置換\test\src"
set EXT="*.vb"

REM 置換前/置換後
set SRC="hoge\r\n|fuga\r\n"
set DST=""

REM https://sakura-editor.github.io/help/HLP000109.html

REM -GREPMODE       Grep実行モードで起動
REM -GKEY=          Grepの検索文字列
REM -GREPR=         Grepの置換文字列
REM -GFILE=         Grepの検索対象のファイル
REM -GFOLDER=       Grepの検索対象のフォルダー
REM -GREPDLG        サクラエディタが起動すると同時にGrepダイアログを表示します。
REM -GCODE=         Grepでの文字コードを指定します。(0=SJIS, 4=UTF-8, 99=自動判別)
REM -GOPT=          Grepの検索条件 [S][L][R][P][W][1|2|3][K][F][B][G][X][C][O][U][H]

REM -GOPT=
REM S               サブフォルダーからも検索
REM L               大文字と小文字を区別
REM R               正規表現
REM P               該当行を出力／未指定時は該当部分だけ出力
REM W               単語単位で探す
REM 1|2|3           結果出力形式。1か2か3のどれかを指定します。(1=ノーマル、2=ファイル毎、3=結果のみ)
REM K               互換性のためだけに残されています。
REM F               ファイル毎最初のみ
REM B               ベースフォルダー表示
REM G               フォルダー毎に表示
REM X               Grep実行後カレントディレクトリを移動しない
REM C               (置換)クリップボードから貼り付け
REM O               (置換)バックアップ作成
REM U               標準出力に出力し、Grep画面にデータを表示しない。コマンドラインからパイプやリダイレクトを指定することで結果を利用できます。
REM H               ヘッダー・フッターを出力しない

REM メイン
%APP% -GREPMODE -GKEY=%SRC% -GREPR=%DST% -GFILE=%EXT% -GFOLDER="%TARGET%" -GCODE=99 -GOPT=SLRU
