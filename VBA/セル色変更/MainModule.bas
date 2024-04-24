Attribute VB_Name = "MainModule"
Option Explicit

'定数
Const MAIN_WS = "main"

Const KEY_FILE_PATH = "FILE_PATH"
Const KEY_INPUT_SHEET_NAME = "INPUT_SHEET_NAME"
Const KEY_WORD = "WORD"
Const KEY_BGCOL = "BGCOL"
Const KEY_LOOKAT = "LOOKAT"
Const KEY_MATCHCASE = "MATCHCASE"
Const KEY_MATCHBYTE = "MATCHBYTE"

Const LOOKAT_WHOLE = "完全一致"
Const LOOKAT_PART = "部分一致"
Const MATCHCASE_TRUE = "大小区別する"
Const MATCHCASE_FALSE = "区別しない"
Const MATCHBYTE_TRUE = "全半角区別する"
Const MATCHBYTE_FALSE = "区別しない"

Const DICT = "Scripting.Dictionary"

Sub ボタン1_Click()

On Error GoTo Exception
    
    Dim map As Object
    Set map = CreateObject(DICT)
    
    Dim searchInfos As Collection
    
    Worksheets(MAIN_WS).Select

    map.Add KEY_FILE_PATH, Range("B5").Value
    map.Add KEY_INPUT_SHEET_NAME, Range("B9").Value
    map.Add KEY_WORD, "B15"
    map.Add KEY_BGCOL, "C15"
    map.Add KEY_LOOKAT, "D15"
    map.Add KEY_MATCHCASE, "E15"
    map.Add KEY_MATCHBYTE, "F15"
    
    Dim obj As SearchInfoDataModel

    '検索情報を取得
    Set searchInfos = GetSearchInfo(map)


    If map(KEY_FILE_PATH) = "" Then
        Main map, searchInfos
        
        Worksheets(MAIN_WS).Select
    Else
        Application.DisplayAlerts = False
        Workbooks.Open map(KEY_FILE_PATH)
        Application.DisplayAlerts = True
                
        Main map, searchInfos
    End If

    '1つも見つからなかった場合は、結果をA列に書き込む
    With ActiveSheet
        For Each obj In searchInfos
            Cells(obj.GetRow, 1).Value = ""
            If obj.GetResult = False Then
                Cells(obj.GetRow, 1).Value = "Not found."
            End If
        Next
    End With

    MsgBox "Success!"
    
    Exit Sub

Exception:
    MsgBox Err.Number & vbCrLf & Err.Description

End Sub

Sub Clear_Click()
    If CommonModule.ShowYesNoMessageBox("検索情報をクリアしますか?") = False Then
        Exit Sub
    End If
    
    ClearRange ("A16")
    ClearRange ("B16")
End Sub

Function GetSearchInfo(ByVal map As Object) As Collection
    Dim i, j, row, Clm As Integer
    Dim wk As String
    Dim ary_wk As Variant
    Dim word_clm, bgcol_clm, lookat_clm, matchcase_clm, matchbyte_clm As Integer
    Dim word_pos, bgcol_pos, lookat_pos, matchcase_pos, matchbyte_pos As String
    Dim word As String
    Dim bgcol As Long
    Dim lookat, matchcase, matchbyte As Boolean
    
    Dim obj As SearchInfoDataModel

    Dim searchInfos As New Collection
    
    word_pos = map(KEY_WORD)
    bgcol_pos = map(KEY_BGCOL)
    lookat_pos = map(KEY_LOOKAT)
    matchcase_pos = map(KEY_MATCHCASE)
    matchbyte_pos = map(KEY_MATCHBYTE)

    With ActiveSheet
        row = .Range(word_pos).row
        word_clm = .Range(word_pos).Column
        bgcol_clm = .Range(bgcol_pos).Column
        lookat_clm = .Range(lookat_pos).Column
        matchcase_clm = .Range(matchcase_pos).Column
        matchbyte_clm = .Range(matchbyte_pos).Column
    
        i = 1
    
        Do
            '---------------
            word = Cells(row + i, word_clm).Value
            
            If word = "" Then
                Exit Do
            End If
            
            '---------------
            bgcol = Cells(row + i, bgcol_clm).Interior.Color
            
            '---------------
            wk = Cells(row + i, lookat_clm).Value
            lookat = False
            If wk = "" Or wk = LOOKAT_WHOLE Then
                lookat = True
            End If
            
            '---------------
            wk = Cells(row + i, matchcase_clm).Value
            matchcase = False
            If wk = "" Or wk = MATCHCASE_TRUE Then
                matchcase = True
            End If
            
            '---------------
            wk = Cells(row + i, matchbyte_clm).Value
            matchbyte = False
            If wk = "" Or wk = MATCHBYTE_TRUE Then
                matchbyte = True
            End If
            
            
            Set obj = New SearchInfoDataModel
            obj.SetNum = i
            obj.SetRow = row + i
            obj.SetWord = word
            obj.SetBgCol = bgcol
            obj.SetLookAt = lookat
            obj.SetMatchCase = matchcase
            obj.SetMatchByte = matchbyte
            obj.SetResult = False
            
            searchInfos.Add obj
            
            i = i + 1
        Loop
    
    End With
    
    Set GetSearchInfo = searchInfos

End Function

Function Main(ByVal map As Object, ByVal searchInfos As Collection)

    Dim si As Variant
    Dim i, k As Long
    Dim obj As SearchInfoDataModel
    Dim obj_sheet As Worksheet
    Dim in_sheet As String
    Dim sheet As Variant
    Dim wk_range As Range
    
    Dim targetSheetNames As New Collection
    
    
    in_sheet = map(KEY_INPUT_SHEET_NAME)
    
    
    If in_sheet = "" Then
        '入力シート名が未指定の場合は全シートを対象にする
        For Each obj_sheet In ThisWorkbook.Worksheets
            targetSheetNames.Add obj_sheet.Name
        Next
    Else
        '入力シート名指定の場合
    
        If IsExistSheet(in_sheet) = False Then
            Err.Raise 1001, "MainModule.Main", "Input sheet is not exist!(" & in_sheet & ")"
        End If
    
        targetSheetNames.Add in_sheet
    End If
    
    
    For Each sheet In targetSheetNames
        'Debug.Print sheet
        Worksheets(sheet).Select
        
        For Each obj In searchInfos
            UpdateCells obj
        Next
    Next
    
End Function


Function UpdateCells(ByRef obj As SearchInfoDataModel)
    'Debug.Print obj.ToString()
    
    Dim foundCell As Range, firstCell As Range, target As Range
    
    Dim word As String
    Dim bgcol As Long
    Dim lookat As Integer
    Dim matchcase, matchbyte As Boolean
        
    word = obj.GetWord()
    bgcol = obj.GetBgCol()
    lookat = obj.GetLookAt()
    matchcase = obj.GetMatchCase()
    matchbyte = obj.GetMatchByte()
    
    
    Set foundCell = Cells.Find(What:=word, lookat:=lookat, matchcase:=matchcase, matchbyte:=matchbyte)
     
    If foundCell Is Nothing Then
        Exit Function
    Else
        Set firstCell = foundCell
        Set target = foundCell
    End If
    
    Do
        Set foundCell = Cells.FindNext(foundCell)
        If foundCell.Address = firstCell.Address Then
            Exit Do
        Else
            Set target = Union(target, foundCell)
        End If
    Loop
    
    target.Interior.Color = bgcol
    obj.SetResult = True
     
End Function


