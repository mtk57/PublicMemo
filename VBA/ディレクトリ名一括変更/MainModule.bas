Attribute VB_Name = "MainModule"
Option Explicit

Sub ボタン1_Click()

On Error GoTo Exception
    
    Dim ret As String

    Worksheets("main").Select
                
    ret = Main(Range("B5").Value)

    MsgBox ret
    
    Exit Sub

Exception:
    MsgBox Err.Number & vbCrLf & Err.Description

End Sub


Function Main(ByVal targetPath As String) As String
    Dim ret As String
    Dim fso As FileSystemObject
    Dim rootDir As Folder
       
    ret = "success!"
    
    If targetPath = "" Then
        ret = "Path is nothing!"
        GoTo FINISH
    End If
    
    If IsExistDir(targetPath) = False Then
        ret = "Path is not exist!"
        GoTo FINISH
    End If
    
    Set fso = New FileSystemObject
    Set rootDir = fso.GetFolder(targetPath)
    
    RenameDir rootDir
    
FINISH:
    Set fso = Nothing
    Main = ret
    
End Function

Function RenameDir(ByVal targetDir As Folder)
    Dim fso As FileSystemObject
    Dim dir As Folder
    Dim dir2 As Folder
    Dim dirName As String
    Dim newDirName As Integer
    Dim i As Long
    
    Set fso = New FileSystemObject

    For Each dir In targetDir.SubFolders
        dirName = dir.Name
        
        'とりあえず1つのディレクトリ内の最大サブディレクトリ数は100,000としておく
        For i = 0 To 100000
            newDirName = i
            
            '現在注目しているディレクトリにnewDirNameは存在するか調べる
            For Each dir2 In targetDir.SubFolders
                If dir2.Name = Str(newDirName) Then
                    '存在するのでnewDirNameを更新する
                    GoTo CONTINUE_I
                End If
            Next dir2
            
            '存在しないのでループを抜ける
            Exit For
CONTINUE_I:
        Next i

        
        'リネーム
        dir.Name = Str(newDirName)
        
        '再帰
        RenameDir dir
    Next dir

    Set fso = Nothing

End Function



