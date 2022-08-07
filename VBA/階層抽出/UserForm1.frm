VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   6293
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   8792.001
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const SHEET_MAIN = "main"
Const INPUT_SHEET_NAME = "Input"
Const START_ROW = 4

Const ROOT_NODE_ID = "B3"

Const PK_PREFIX = "PK_"
Const PK_CLM_NUM = 1
Const NODE_PARENT_CLM_NUM = 2
Const NODE_CHILID_CLM_NUM = 3


Private Sub UserForm_Initialize()
On Error GoTo Exception
    
    Dim i, rowCnt As Long
    
    Dim WS As Worksheet
    Dim TV As TreeView
    
    Dim rootNodeId As String
    Dim parentNodeKey As String
    Dim nodeKey As String
    Dim nodeText As String
    
    rootNodeId = Worksheets(SHEET_MAIN).Range(ROOT_NODE_ID).Value
    
    Worksheets(INPUT_SHEET_NAME).Select
    
    Set WS = Worksheets(INPUT_SHEET_NAME)
    Set TV = Me.TreeView1
    
    Me.TreeView1.Nodes.Clear
    
    With WS
        'ルートノードを登録
        TV.Nodes.Add _
            Relative:=Null, _
            Relationship:=0, _
            key:=PK_PREFIX & rootNodeId, _
            Text:=rootNodeId
        
        'ルートノード配下のノードを登録
        rowCnt = .Cells(.Rows.Count, NODE_PARENT_CLM_NUM).End(xlUp).row
        
        For i = START_ROW To rowCnt
        
            parentNodeKey = PK_PREFIX & .Cells(i, NODE_PARENT_CLM_NUM)
            nodeKey = PK_PREFIX & .Cells(i, NODE_CHILID_CLM_NUM)
            nodeText = Cells(i, NODE_CHILID_CLM_NUM)

            TV.Nodes.Add _
                Relative:=parentNodeKey, _
                Relationship:=tvwChild, _
                key:=nodeKey, _
                Text:=nodeText

        Next i
    End With
    
    GoTo Finally
    
Exception:
    MsgBox Err.Number & vbCrLf & Err.Description
    
Finally:
    Set WS = Nothing
    Set TV = Nothing
    
    Worksheets(SHEET_MAIN).Select
End Sub
