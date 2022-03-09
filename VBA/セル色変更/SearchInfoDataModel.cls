VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SearchInfoDataModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private num_ As Long
Private row_ As Long
Private word_ As String
Private bgcol_ As Long
Private lookat_ As Integer
Private matchcase_ As Boolean
Private matchbyte_ As Boolean
Private result_ As Boolean

Public Function ToString() As String
    ToString = "[" & num_ & "]:row_:" & row_ & " word_:" & word_ & " bgcol_:" & bgcol_ & " lookat_:" & lookat_ & " matchcase_:" & matchcase_ & " matchbyte_:" & matchbyte_
End Function

Public Property Let SetNum(ByVal num As Long)
    num_ = num
End Property

Public Property Get GetNum() As Long
    GetNum = num_
End Property

Public Property Let SetRow(ByVal row As Long)
    row_ = row
End Property

Public Property Get GetRow() As Long
    GetRow = row_
End Property

Public Property Let SetWord(ByVal word As String)
    word_ = word
End Property

Public Property Get GetWord() As String
    GetWord = word_
End Property

Public Property Let SetBgCol(ByVal bgcol As Long)
    bgcol_ = bgcol
End Property

Public Property Get GetBgCol() As Long
    GetBgCol = bgcol_
End Property

Public Property Let SetLookAt(ByVal lookat As Boolean)
    If lookat = True Then
        lookat_ = xlWhole
    Else
        lookat_ = xlPart
    End If
End Property

Public Property Get GetLookAt() As Integer
    GetLookAt = lookat_
End Property

Public Property Let SetMatchCase(ByVal matchcase As Boolean)
    matchcase_ = matchcase
End Property

Public Property Get GetMatchCase() As Boolean
    GetMatchCase = matchcase_
End Property

Public Property Let SetMatchByte(ByVal matchbyte As Boolean)
    matchbyte_ = matchbyte
End Property

Public Property Get GetMatchByte() As Boolean
    GetMatchByte = matchbyte_
End Property

Public Property Let SetResult(ByVal result As Boolean)
    result_ = result
End Property

Public Property Get GetResult() As Boolean
    GetResult = result_
End Property
