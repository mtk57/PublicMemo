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

Private is_enable As String
Private type_vb6 As String
Private property_vb6 As String
Private type_vbnet As String
Private property_vbnet As String
Private value_vbnet As String
Private is_apply_twips_per_pixel As bool
Private is_apply_twips_per_pixel_s As String

Private Const CLM_ = "L"
Private Const ROW_ = 7

'コンストラクタ
Private Sub Class_Initialize()
    is_enable = ""
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

    is_enable = sub_sheet.Range("B" & row).value
    type_vb6 = sub_sheet.Range("C" & row).value
    property_vb6 = sub_sheet.Range("D" & row).value
    type_vbnet = sub_sheet.Range("E" & row).value
    property_vbnet = sub_sheet.Range("F" & row).value
    value_vbnet = sub_sheet.Range("G" & row).value
    is_apply_twips_per_pixel_s = sub_sheet.Range("H" & row).value
End Sub

'検証
Public Sub Validate()
    If is_enable = "" Or _
       is_enable = "Stopper" Then
       is_enable = "Stopper"
    ElseIf is_enable = "Enable" Then
        is_enable = "Enable"
    Else
        is_enable = "Disable"
    End If

    If is_enable = "Disable" Or is_enable = "Stopper" Then
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
Public Function GetIsEnable() As String
    GetIsEnable = is_enable
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

'=========================================
'Setter
'=========================================
Public Sub SetIsEnable(ByVal val As String)
    is_enable = val
End Sub

Public Sub SetTypeVB6(ByVal val As String)
    type_vb6 = val
End Sub

Public Sub SetPropertyVB6(ByVal val As String)
    property_vb6 = val
End Sub

Public Sub SetTypeVBdotNet(ByVal val As String)
    type_vbnet = val
End Sub

Public Sub SetPropertyVBdotNet(ByVal val As String)
    property_vbnet = val
End Sub

Public Sub SetValueVBdotNet(ByVal val As String)
    value_vbnet = val
End Sub

Public Sub SetIsApplyTwipsPerPixel(ByVal val As Boolean)
    is_apply_twips_per_pixel = val
End Sub
