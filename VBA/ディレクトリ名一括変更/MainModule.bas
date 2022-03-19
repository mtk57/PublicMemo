Attribute VB_Name = "MainModule"
Option Explicit

Sub �{�^��1_Click()

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
        
        '�Ƃ肠����1�̃f�B���N�g�����̍ő�T�u�f�B���N�g������100,000�Ƃ��Ă���
        For i = 0 To 100000
            newDirName = i
            
            '���ݒ��ڂ��Ă���f�B���N�g����newDirName�͑��݂��邩���ׂ�
            For Each dir2 In targetDir.SubFolders
                If dir2.Name = Str(newDirName) Then
                    '���݂���̂�newDirName���X�V����
                    GoTo CONTINUE_I
                End If
            Next dir2
            
            '���݂��Ȃ��̂Ń��[�v�𔲂���
            Exit For
CONTINUE_I:
        Next i

        
        '���l�[��
        dir.Name = Str(newDirName)
        
        '�ċA
        RenameDir dir
    Next dir

    Set fso = Nothing

End Function



