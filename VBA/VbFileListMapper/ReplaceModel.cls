VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ReplaceModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_CLMS = 3
Private Const CLM_SRC_PRJ = 0
Private Const CLM_SRC_REF = 1
Private Const CLM_DST_PRJ = 2
Private Const CLM_DST_REF = 3

'メンバー変数
Private ary_() As String
Private now_row_cnt_ As Long
Private max_row_cnt_ As Long
Private is_empty As Boolean

'コンストラクタ
Private Sub Class_Initialize()
    now_row_cnt_ = 0
    max_row_cnt_ = -1
    is_empty = True
End Sub

'初期化
Public Sub Init(ByVal row_cnt As Long)
    If row_cnt < 0 Then
        Exit Sub
    End If
    
    ReDim Preserve ary_(row_cnt, MAX_CLMS)
    max_row_cnt_ = row_cnt
    is_empty = False
End Sub

'=========================================
'Setter
'=========================================
Public Sub Append( _
    ByVal src_prj_path As String, _
    ByVal src_ref_path As String, _
    ByVal dst_prj_path As String, _
    ByVal dst_ref_path As String _
)
    ary_(now_row_cnt_, CLM_SRC_PRJ) = src_prj_path
    ary_(now_row_cnt_, CLM_SRC_REF) = src_ref_path
    ary_(now_row_cnt_, CLM_DST_PRJ) = dst_prj_path
    ary_(now_row_cnt_, CLM_DST_REF) = dst_ref_path
    
    now_row_cnt_ = now_row_cnt_ + 1
End Sub



'=========================================
'Getter
'=========================================
Public Function IsEmpty() As Boolean
    IsEmpty = is_empty
End Function

Public Function GetRowCount() As Long
    GetRowCount = UBound(ary_, 1)
End Function

Public Function GetClmCount() As Long
    GetClmCount = UBound(ary_, 2)
End Function

Public Function GetSrcPrjPath(ByVal row_num As Long) As String
    GetSrcPrjPath = ary_(row_num, CLM_SRC_PRJ)
End Function

Public Function GetSrcRefPath(ByVal row_num As Long) As String
    GetSrcRefPath = ary_(row_num, CLM_SRC_REF)
End Function

Public Function GetDstPrjPath(ByVal row_num As Long) As String
    GetDstPrjPath = ary_(row_num, CLM_DST_PRJ)
End Function

Public Function GetDstRefPath(ByVal row_num As Long) As String
    GetDstRefPath = ary_(row_num, CLM_DST_REF)
End Function

Public Function GetIndexSrcPrjPath(ByVal path As String) As Long
    GetIndexSrcPrjPath = -1
    
    Dim i As Long
    
    For i = 0 To GetRowCount()
        If GetSrcPrjPath(i) = path Then
            GetIndexSrcPrjPath = i
            Exit Function
        End If
    Next i
End Function

Public Function GetIndexSrcRefPath(ByVal path As String) As Long
    GetIndexSrcRefPath = -1
    
    Dim i As Long
    
    For i = 0 To GetRowCount()
        If GetSrcRefPath(i) = path Then
            GetIndexSrcRefPath = i
            Exit Function
        End If
    Next i
End Function

'=========================================
'Public I/F
'=========================================

