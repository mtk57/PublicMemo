Attribute VB_Name = "Common"
Option Explicit

'���O�t�@�C���ԍ�
Private logfile_num As Integer
Private is_log_opened As Boolean

'-------------------------------------------------------------
'���O�t�@�C�����I�[�v������
' logfile_path : IN : ���O�t�@�C���p�X(��΃p�X)
'-------------------------------------------------------------
Public Sub OpenLog(ByVal logfile_path As String)
    If is_log_opened = True Then
        '���łɃI�[�v�����Ă���̂Ŗ���
        Exit Sub
    End If
    logfile_num = FreeFile()
    Open logfile_path For Append As logfile_num
    is_log_opened = True
End Sub

'-------------------------------------------------------------
'���O�t�@�C���ɏ�������
' contents : IN : �������ޓ��e
'-------------------------------------------------------------
Public Sub WriteLog(ByVal contents As String)
    If is_log_opened = False Then
        '�I�[�v������Ă��Ȃ��̂Ŗ���
        Exit Sub
    End If
    Print #logfile_num, Format(Date, "yyyy/mm/dd") & " " & Format(Now, "hh:mm:ss") & ":" & contents
End Sub

'-------------------------------------------------------------
'���O�t�@�C�����N���[�Y����
'-------------------------------------------------------------
Public Sub CloseLog()
    If is_log_opened = False Then
        '�I�[�v������Ă��Ȃ��̂Ŗ���
        Exit Sub
    End If
    Close logfile_num
    logfile_num = -1
    is_log_opened = False
End Sub

'-------------------------------------------------------------
'���O�t�@�C���ɏ�������
' contents : IN : �������ޓ��e
'-------------------------------------------------------------
Public Sub WriteLogSimple(ByVal contents As String)
    Dim file_num As Integer
    file_num = FreeFile()
    Open "Logger.log" For Append As file_num
    Print #file_num, Format(Date, "yyyy/mm/dd") & " " & Format(Now, "hh:mm:ss") & ":" & contents
    Close file_num
    file_num = -1
End Sub

