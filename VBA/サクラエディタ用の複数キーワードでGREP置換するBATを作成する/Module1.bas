Attribute VB_Name = "Module1"
Option Explicit

Sub ボタン_Click()
    
On Error GoTo Exception
    
    Dim row, clm_offset As Integer
    Dim cell_value As String
    Dim command As String
    
    Const START_ROW = 11
    Const START_CLM = 2
    
    Worksheets("main").Select
    
    With ActiveSheet
    
        row = START_ROW
        
        '行ループ
        Do
            
            '================================================
            '■共通設定
            '------------------------------------------------
            '1.サクラエディタのパス
            command = Range("G5").Value
        
        
            '================================================
            '■キーワード毎の設定
            '------------------------------------------------
            '1.-GOPT=
            command = command & " -GOPT="
            
            '1-1.S
            clm_offset = 0
            cell_value = Cells(row, START_CLM + clm_offset).Value
            If cell_value <> "" Then
                command = command & "S"
            End If
            
            '1-2.L
            clm_offset = clm_offset + 1
            cell_value = Cells(row, START_CLM + clm_offset).Value
            If cell_value <> "" Then
                command = command & "L"
            End If
            
            '1-3.R
            clm_offset = clm_offset + 1
            cell_value = Cells(row, START_CLM + clm_offset).Value
            If cell_value <> "" Then
                command = command & "R"
            End If
            
            '1-4.P
            clm_offset = clm_offset + 1
            cell_value = Cells(row, START_CLM + clm_offset).Value
            If cell_value <> "" Then
                command = command & "P"
            End If
            
            '1-5.W
            clm_offset = clm_offset + 1
            cell_value = Cells(row, START_CLM + clm_offset).Value
            If cell_value <> "" Then
                command = command & "W"
            End If
            
            '1-6.1 or 2 or 3
            clm_offset = clm_offset + 1
            cell_value = Cells(row, START_CLM + clm_offset).Value
            If cell_value <> "" Then
                command = command & cell_value
            End If
            
            '1-7.F
            clm_offset = clm_offset + 1
            cell_value = Cells(row, START_CLM + clm_offset).Value
            If cell_value <> "" Then
                command = command & "F"
            End If
            
            '1-8.B
            clm_offset = clm_offset + 1
            cell_value = Cells(row, START_CLM + clm_offset).Value
            If cell_value <> "" Then
                command = command & "B"
            End If
            
            '1-9.G
            clm_offset = clm_offset + 1
            cell_value = Cells(row, START_CLM + clm_offset).Value
            If cell_value <> "" Then
                command = command & "G"
            End If
            
            '1-10.X
            clm_offset = clm_offset + 1
            cell_value = Cells(row, START_CLM + clm_offset).Value
            If cell_value <> "" Then
                command = command & "X"
            End If
            
            '1-11.C
            clm_offset = clm_offset + 1
            cell_value = Cells(row, START_CLM + clm_offset).Value
            If cell_value <> "" Then
                command = command & "C"
            End If
            
            '1-12.O
            clm_offset = clm_offset + 1
            cell_value = Cells(row, START_CLM + clm_offset).Value
            If cell_value <> "" Then
                command = command & "O"
            End If
            
            '1-13.U
            clm_offset = clm_offset + 1
            cell_value = Cells(row, START_CLM + clm_offset).Value
            If cell_value <> "" Then
                command = command & "U"
            End If
            
            '1-14.H
            clm_offset = clm_offset + 1
            cell_value = Cells(row, START_CLM + clm_offset).Value
            If cell_value <> "" Then
                command = command & "H"
            End If

            '------------------------------------------------
            '2.-GREPMODE
            clm_offset = clm_offset + 1
            cell_value = Cells(row, START_CLM + clm_offset).Value
            If cell_value = "" Then
                'GREPMODEが空なので終了
                GoTo FINISH
            End If
            command = command & " -GREPMODE"
            
            '------------------------------------------------
            '3.-GKEY=
            clm_offset = clm_offset + 1
            cell_value = Cells(row, START_CLM + clm_offset).Value
            If cell_value = "" Then
                '必須項目が空なので次の行へ
                GoTo CONTINUE
            End If
            command = command & " -GKEY=" & cell_value
            
            '------------------------------------------------
            '4.-GREPR=
            clm_offset = clm_offset + 1
            cell_value = Cells(row, START_CLM + clm_offset).Value
            If cell_value <> "" Then
                command = command & " -GREPR=" & cell_value
            End If
        
            '------------------------------------------------
            '5.-GFILE=
            clm_offset = clm_offset + 1
            cell_value = Cells(row, START_CLM + clm_offset).Value
            If cell_value = "" Then
                '必須項目が空なので次の行へ
                GoTo CONTINUE
            End If
            command = command & " -GFILE=" & cell_value
            
            '------------------------------------------------
            '6.-GFOLDER=
            clm_offset = clm_offset + 1
            cell_value = Cells(row, START_CLM + clm_offset).Value
            If cell_value = "" Then
                '必須項目が空なので次の行へ
                GoTo CONTINUE
            End If
            command = command & " -GFOLDER=" & cell_value
            
            '------------------------------------------------
            '7.-GREPDLG
            clm_offset = clm_offset + 1
            cell_value = Cells(row, START_CLM + clm_offset).Value
            If cell_value <> "" Then
                command = command & " -GREPDLG"
            End If
        
            '------------------------------------------------
            '8.-GCODE=
            clm_offset = clm_offset + 1
            cell_value = Cells(row, START_CLM + clm_offset).Value
            If cell_value <> "" Then
                command = command & " -GCODE=" & cell_value
            End If

            '================================================
            'コマンド列に出力
            clm_offset = clm_offset + 1
            Debug.Print command
            Cells(row, START_CLM + clm_offset).Value = command
        
CONTINUE:
            row = row + 1
        Loop
        
    End With
    
    
FINISH:
    MsgBox "Success!"
    
    Exit Sub
    
Exception:
    MsgBox Err.Number & vbCrLf & Err.Description
    
End Sub
