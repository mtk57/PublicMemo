echo off
REM 注意
REM 以下のps1スクリプトを実行する前に環境変数にFTP接続情報をセットしないと失敗する
powershell.exe -ExecutionPolicy Bypass -File "C:\_git\PublicMemo\VBA\FtpTest\ps1test\ftp_script.ps1"
pause
