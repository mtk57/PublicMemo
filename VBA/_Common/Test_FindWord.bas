Attribute VB_Name = "Test_FindWord"
Option Explicit

Public Sub TestFindWord()
    Dim testsPassed As Integer
    Dim totalTests As Integer
    
    totalTests = 0
    testsPassed = 0
    
    ' �e�X�g�P�[�X1: �啶������������ʂ��A������v�Ō���
    totalTests = totalTests + 1
    If TestCase("Hello World", "World", True, False, False, True) Then
        testsPassed = testsPassed + 1
    End If
    
    ' �e�X�g�P�[�X2: �啶������������ʂ����A���S��v�Ō���
    totalTests = totalTests + 1
    If TestCase("Hello World", "hello world", False, True, False, True) Then
        testsPassed = testsPassed + 1
    End If
    
    ' �e�X�g�P�[�X3: ���K�\�����g�p���Đ���������
    totalTests = totalTests + 1
    If TestCase("abc123def", "\d+", False, False, True, True) Then
        testsPassed = testsPassed + 1
    End If
    
    ' �e�X�g�P�[�X4: �啶������������ʂ��A���S��v�Ō����i���s����P�[�X�j
    totalTests = totalTests + 1
    If TestCase("Hello World", "World", True, True, False, False) Then
        testsPassed = testsPassed + 1
    End If
    
    ' �e�X�g�P�[�X5: ��̕����������
    totalTests = totalTests + 1
    If TestCase("Hello World", "", True, False, False, True) Then
        testsPassed = testsPassed + 1
    End If
    
    ' �e�X�g�P�[�X6: ���K�\���ŕ�����̐擪�Ɩ������w��
    totalTests = totalTests + 1
    If TestCase("Hello World", "^Hello World$", False, False, True, True) Then
        testsPassed = testsPassed + 1
    End If
    
    ' �e�X�g�P�[�X7: ���݂��Ȃ������������
    totalTests = totalTests + 1
    If TestCase("Hello World", "Goodbye", False, False, False, False) Then
        testsPassed = testsPassed + 1
    End If
    
    ' ���ʂ̏o��
    Debug.Print "�e�X�g����: " & testsPassed & " / " & totalTests & " �p�X"
    If testsPassed = totalTests Then
        MsgBox "���ׂẴe�X�g�Ƀp�X���܂����I"
    Else
        MsgBox "���s�����e�X�g������܂��B��L�̏ڍׂ��m�F���Ă��������B"
    End If
End Sub

Private Function TestCase(targetStr As String, findStr As String, letterCase As Boolean, exactMatch As Boolean, useRegEx As Boolean, expectedResult As Boolean) As Boolean
    Dim result As Boolean
    result = FindWord(targetStr, findStr, letterCase, exactMatch, useRegEx)
    
    Debug.Print "�e�X�g�P�[�X: " & _
                "targetStr='" & targetStr & "', " & _
                "findStr='" & findStr & "', " & _
                "letterCase=" & letterCase & ", " & _
                "exactMatch=" & exactMatch & ", " & _
                "useRegEx=" & useRegEx
    Debug.Print "  ���Ҍ���: " & expectedResult & ", ���ۂ̌���: " & result
    
    If result = expectedResult Then
        Debug.Print "  �e�X�g����"
        TestCase = True
    Else
        Debug.Print "  �e�X�g���s"
        TestCase = False
    End If
    
    Debug.Print ""
End Function

