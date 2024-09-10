VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParseResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'メンバー変数
Private target As String
Private format_type As String
Private file_path As String
Private extension As String
Private row_num As Long
Private clm_num As Long
Private encode As String
Private raw_line As String
Private with_codes() As String
Private with_class As String
Private with_members() As String

'コンストラクタ
Private Sub Class_Initialize()
    target = ""
    format_type = ""
    file_path = ""
    extension = ""
    row_num = 0
    clm_num = 0
    encode = ""
    raw_line = ""
    with_class = ""
End Sub

'初期化
Public Sub Init(ByVal target_ As String, ByVal format_type_ As String)
    target = target_
    format_type = format_type_
    
    Parse
End Sub

Private Sub Parse()
    If format_type = "sakura" Then
        ParseForSakura
    End If
End Sub

'targetに想定する文字列は以下とする
'"C:\_git\PublicMemo\VBA\WithParser\testdata\src\with_sample.bas(2,4)  [UTF-8]:     With MyObject"
'
'つまり、以下の構成であること。
'ファイルの絶対パス        :"C:\_git\PublicMemo\VBA\WithParser\testdata\src\with_sample.bas"
'(行番号, 列番号)          :"(2,4)"
'半角スペース*2            :"  "
'エンコード(例:[UTF-8])    :"[UTF-8]"
'半角コロン                :":"
'GREP結果                  :"     With MyObject""
Private Sub ParseForSakura()
    Dim wk() As String
    Dim path_wk As String
    Dim pos_wk As String
    Dim pos_list() As String
    Dim i As Long
    
    'いったん半角SPで分解
    wk = Split(target, " ")
    wk = Common.DeleteEmptyArray(wk)

    'ファイルパスを取得
    path_wk = wk(0)
    file_path = Mid(path_wk, 1, InStrRev(path_wk, "(") - 1)
    
    '拡張子を取得
    extension = Common.GetFileExtension(file_path)
    
    'いったん"("から")"までを取得
    pos_wk = Mid(path_wk, InStrRev(path_wk, "("), Len(path_wk))
    
    '行・列を取得
    pos_list = Common.DeleteEmptyArray(Common.GetMatchByRegExp(pos_wk, "\d*", True))
    If UBound(pos_list) < 1 Then
        Err.Raise 53, , "GREP結果の行列フォーマットが不正です (target=" & target & ")"
    End If
    row_num = pos_list(0)
    clm_num = pos_list(1)
    
    'エンコードを取得
    For i = 0 To UBound(wk)
        If Common.IsMatchByRegExp(wk(i), "\[.*\]", True) = True Then
            encode = Replace(Replace(wk(i), "[", ""), "]:", "")
            Exit For
        End If
    Next i
    
    If encode = "" Then
        Err.Raise 53, , "GREP結果からエンコード文字列が見つかりません (target=" & target & ")"
    ElseIf encode <> "SJIS" And encode <> "UTF-8" Then
        Err.Raise 53, , "未サポートのエンコードです (target=" & target & ")"
    End If
    
    'GREP結果を取得
    raw_line = Mid(target, InStr(target, "]:") + 2, Len(target))
    
End Sub

'--------------------
'Getter
'--------------------

Public Function GetTarget() As String
    GetTarget = target
End Function

Public Function GetFormatType() As String
    GetFormatType = format_type
End Function

Public Function GetFilePath() As String
    GetFilePath = file_path
End Function

Public Function GetExtension() As String
    GetExtension = extension
End Function

Public Function GetRowNum() As Long
    GetRowNum = row_num
End Function

Public Function GetClmNum() As Long
    GetClmNum = clm_num
End Function

Public Function GetEncode() As String
    GetEncode = encode
End Function

Public Function GetRawLine() As String
    GetRawLine = raw_line
End Function

Public Function GetWithCodes() As String()
    GetWithCodes = with_codes
End Function

Public Function GetWithCodesCount() As Long
    If Common.IsEmptyArrayLong(with_codes) = True Then
        GetWithCodesCount = 0
        Exit Function
    End If
    
    Dim cnt As Long: cnt = UBound(with_codes)
    cnt = cnt + 1
    GetWithCodesCount = cnt
End Function

Public Function GetWithClass() As String
    GetWithClass = with_class
End Function

Public Function GetWithMembers() As String()
    GetWithMembers = with_members
End Function

Public Function GetWithMembersCount() As Long
    If Common.IsEmptyArrayLong(with_members) = True Then
        GetWithMembersCount = 0
        Exit Function
    End If
    
    Dim cnt As Long: cnt = UBound(with_members)
    cnt = cnt + 1
    GetWithMembersCount = cnt
End Function

'--------------------
'Setter
'--------------------
Public Sub SetWithCodes(ByRef codes() As String)
    with_codes = codes
End Sub

Public Sub SetWithClass(ByVal class_name As String)
    with_class = class_name
End Sub

Public Sub SetWithMembers(ByRef members() As String)
    with_members = members
End Sub

