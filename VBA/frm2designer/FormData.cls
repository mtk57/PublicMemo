VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private root_dict As Object

Private Sub Class_Initialize()
    Set root_dict = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Class_Terminate()
    Set root_dict = Nothing
End Sub

Public Sub Init(ByVal path As String)
    
End Sub

Public Function GetRoot() As Object
    GetRoot = root_dict
End Function

