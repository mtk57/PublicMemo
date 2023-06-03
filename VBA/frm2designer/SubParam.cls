VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SubParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'メンバー変数

Private enable As String
Private type_vb6 As String
Private property_vb6 As String
Private type_vbnet As String
Private property_vbnet As String
Private value_vbnet As String
Private is_apply_twips_per_pixel As Boolean
Private is_apply_twips_per_pixel_s As String

Private Const CLM_ = "L"
Private Const ROW_ = 7

'コンストラクタ
Private Sub Class_Initialize()
    enable = ""
    type_vb6 = ""
    property_vb6 = ""
    type_vbnet = ""
    property_vbnet = ""
    value_vbnet = ""
    is_apply_twips_per_pixel = False
End Sub

'初期化
Public Sub Init(ByVal row As Long)
    Dim sub_sheet As Worksheet
    Set sub_sheet = ThisWorkbook.Sheets("sub")

    enable = sub_sheet.Range("B" & row).value
    type_vb6 = sub_sheet.Range("C" & row).value
    property_vb6 = sub_sheet.Range("D" & row).value
    type_vbnet = sub_sheet.Range("E" & row).value
    property_vbnet = sub_sheet.Range("F" & row).value
    value_vbnet = sub_sheet.Range("G" & row).value
    is_apply_twips_per_pixel_s = sub_sheet.Range("H" & row).value
End Sub

'検証
Public Sub Validate()
    If enable = "" Or _
       enable = "Stopper" Then
       enable = "Stopper"
    ElseIf enable = "Enable" Then
        enable = "Enable"
    Else
        enable = "Disable"
    End If

    If enable = "Disable" Or enable = "Stopper" Then
        Exit Sub
    End If

    If type_vb6 = "" Then
        Err.Raise 53, , "VB6のTypeが指定されていません"
    End If

    If property_vb6 = "" Then
        Err.Raise 53, , "VB6のPropertyが指定されていません"
    End If

    If type_vbnet = "" Then
        Err.Raise 53, , "VB.NETのTypeが指定されていません"
    End If

    If property_vbnet = "" Then
        Err.Raise 53, , "VB.NETのPropertyが指定されていません"
    End If

    If value_vbnet = "" Then
        Err.Raise 53, , "VB.NETのValueが指定されていません"
    End If

    If is_apply_twips_per_pixel_s = "" Or _
       is_apply_twips_per_pixel_s = "NO" Then
       is_apply_twips_per_pixel = False
    Else
        is_apply_twips_per_pixel = True
    End If

End Sub

'=========================================
'Getter
'=========================================
Public Function GetEnable() As String
    GetEnable = enable
End Function

Public Function GetTypeVB6() As String
    GetTypeVB6 = type_vb6
End Function

Public Function GetPropertyVB6() As String
    GetPropertyVB6 = property_vb6
End Function

Public Function GetTypeVBdotNet() As String
    GetTypeVBdotNet = type_vbnet
End Function

Public Function GetPropertyVBdotNet() As String
    GetPropertyVBdotNet = property_vbnet
End Function

Public Function GetValueVBdotNet() As String
    GetValueVBdotNet = value_vbnet
End Function

Public Function GetIsApplyTwipsPerPixel() As Boolean
    GetIsApplyTwipsPerPixel = is_apply_twips_per_pixel
End Function

Public Function GetAllValue() As String
    Dim ret As String
    
    ret = "[Sub Param] " & _
          "enable=(" & enable & "), " & _
          "type_vb6=(" & type_vb6 & "), " & _
          "property_vb6=(" & property_vb6 & "), " & _
          "type_vbnet=(" & type_vbnet & "), " & _
          "property_vbnet=(" & property_vbnet & "), " & _
          "value_vbnet=(" & value_vbnet & "), " & _
          "is_apply_twips_per_pixel=(" & is_apply_twips_per_pixel_s & ")"

    GetAllValue = ret
End Function
