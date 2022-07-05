# ファイルを特定の行数で分割して出力する
# 分割後のファイル名は以下となる。
#  例:hoge.txt が2分割される場合
#     hoge_00000.txt
#     hoge_00001.txt

using namespace System.IO;

#分割行数 ★
$DIV = 10

#分割対象のファイルパス ★
$SRC = "C:\_tmp\test.csv"

#分割ファイルを出力するフォルダパス ★
$DST_DIR = "C:\_git"

#===============================================

#分割ファイル名(拡張子なし)
$DST_FN = [Path]::GetFileNameWithoutExtension($SRC)

#分割ファイル拡張子
$DST_EXT = [Path]::GetExtension($SRC)

$i = 0
$fn = ""
$num = ""

# 分割対象ファイルの中身を指定した行数分読み込んで、次の処理に渡す
Get-Content $SRC -ReadCount $DIV | `
ForEach-Object `
{
    # 出力するファイルパスを作成する
    $num = "{0:D5}" -f $i   #とりあえず5桁0埋め
    $fn = $DST_DIR + "\" + $DST_FN + "_" + $num + $DST_EXT

    # $_には読み込んだ行数分の内容が入っているので、それをOut-Fileに渡す
    $_ | `

    # ファイルに出力する
    Out-File $fn -Encoding UTF8

    $i++
}