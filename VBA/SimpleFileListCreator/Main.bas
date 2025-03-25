Attribute VB_Name = "Main"
Option Explicit

' �t�H���_�z���̑S�t�@�C������Excel�V�[�g�Ɏ擾����}�N��
Sub GetAllFileInfo()
    Dim folderPath As String
    Dim outputSheet As Worksheet
    Dim fso As Object
    Dim folder As Object
    Dim row As Long
    
    ' FileSystemObject���쐬
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' �t�H���_�I���_�C�A���O��\��
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "�t�@�C�������擾����t�H���_��I�����Ă�������"
        .AllowMultiSelect = False
        If .Show = -1 Then
            folderPath = .SelectedItems(1)
        Else
            MsgBox "�t�H���_���I������܂���ł����B", vbExclamation
            Exit Sub
        End If
    End With
    
    ' �t�H���_�����݂��邩�m�F
    If Not fso.FolderExists(folderPath) Then
        MsgBox "�w�肳�ꂽ�t�H���_��������܂���: " & folderPath, vbExclamation
        Exit Sub
    End If
    
    '�V�[�g��ǉ�
    Dim sheet_name_ As String: sheet_name_ = Common.GetNowTimeString()
    Common.AddSheet ActiveWorkbook, sheet_name_
    
    ' �o�͐�V�[�g�̏���
    Application.ScreenUpdating = False
    
    ' �����̃V�[�g������΍폜���A�V�����V�[�g���쐬
    On Error Resume Next
    Application.DisplayAlerts = False
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set outputSheet = ThisWorkbook.Sheets(sheet_name_)
    
    ' �w�b�_�[�̐ݒ�
    outputSheet.Cells(1, 1).value = "�t�@�C���p�X"
    outputSheet.Cells(1, 2).value = "�X�V�N����"
    outputSheet.Cells(1, 3).value = "�X�V�����b"
    outputSheet.Cells(1, 4).value = "�t�@�C���T�C�Y"
    
    ' �w�b�_�[�s�̏����ݒ�͏ȗ�

    
    ' �����s�ԍ�
    row = 2
    
    ' �t�H���_���擾
    Set folder = fso.GetFolder(folderPath)
    
    ' �t�H���_���̃t�@�C���������W
    ProcessFiles folder, outputSheet, row
    
    ' �񕝂̎��������̂ݎ��{
    outputSheet.Columns("A:D").AutoFit
    
    Application.ScreenUpdating = True
    
    MsgBox "�t�@�C�����̎擾���������܂����B" & vbCrLf & _
           "�t�@�C����: " & (row - 2), vbInformation
End Sub

' �t�H���_���̃t�@�C�����ċA�I�ɏ�������֐�
Private Sub ProcessFiles(folder As Object, outputSheet As Worksheet, ByRef row As Long)
    Dim file As Object
    Dim subfolder As Object
    Dim lastModified As Date
    
    ' �t�H���_���̂��ׂẴt�@�C��������
    For Each file In folder.Files
        ' �t�@�C�������擾
        lastModified = file.DateLastModified
        
        ' �V�[�g�ɏo��
        outputSheet.Cells(row, 1).value = file.path
        outputSheet.Cells(row, 2).value = Format(lastModified, "yyyy/mm/dd")
        outputSheet.Cells(row, 3).value = Format(lastModified, "hh:mm:ss")
        outputSheet.Cells(row, 4).value = file.size
        
        ' �s�ԍ��𑝂₷
        row = row + 1
    Next file
    
    ' �T�u�t�H���_���ċA�I�ɏ���
    For Each subfolder In folder.SubFolders
        ProcessFiles subfolder, outputSheet, row
    Next subfolder
End Sub




