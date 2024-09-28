Attribute VB_Name = "Test_FindWord"
Option Explicit

Public Sub TestFindWord()
    Dim testsPassed As Integer
    Dim totalTests As Integer
    
    totalTests = 0
    testsPassed = 0
    
    ' テストケース1: 大文字小文字を区別し、部分一致で検索
    totalTests = totalTests + 1
    If TestCase("Hello World", "World", True, False, False, True) Then
        testsPassed = testsPassed + 1
    End If
    
    ' テストケース2: 大文字小文字を区別せず、完全一致で検索
    totalTests = totalTests + 1
    If TestCase("Hello World", "hello world", False, True, False, True) Then
        testsPassed = testsPassed + 1
    End If
    
    ' テストケース3: 正規表現を使用して数字を検索
    totalTests = totalTests + 1
    If TestCase("abc123def", "\d+", False, False, True, True) Then
        testsPassed = testsPassed + 1
    End If
    
    ' テストケース4: 大文字小文字を区別し、完全一致で検索（失敗するケース）
    totalTests = totalTests + 1
    If TestCase("Hello World", "World", True, True, False, False) Then
        testsPassed = testsPassed + 1
    End If
    
    ' テストケース5: 空の文字列を検索
    totalTests = totalTests + 1
    If TestCase("Hello World", "", True, False, False, True) Then
        testsPassed = testsPassed + 1
    End If
    
    ' テストケース6: 正規表現で文字列の先頭と末尾を指定
    totalTests = totalTests + 1
    If TestCase("Hello World", "^Hello World$", False, False, True, True) Then
        testsPassed = testsPassed + 1
    End If
    
    ' テストケース7: 存在しない文字列を検索
    totalTests = totalTests + 1
    If TestCase("Hello World", "Goodbye", False, False, False, False) Then
        testsPassed = testsPassed + 1
    End If
    
    ' 結果の出力
    Debug.Print "テスト結果: " & testsPassed & " / " & totalTests & " パス"
    If testsPassed = totalTests Then
        MsgBox "すべてのテストにパスしました！"
    Else
        MsgBox "失敗したテストがあります。上記の詳細を確認してください。"
    End If
End Sub

Private Function TestCase(targetStr As String, findStr As String, letterCase As Boolean, exactMatch As Boolean, useRegEx As Boolean, expectedResult As Boolean) As Boolean
    Dim result As Boolean
    result = FindWord(targetStr, findStr, letterCase, exactMatch, useRegEx)
    
    Debug.Print "テストケース: " & _
                "targetStr='" & targetStr & "', " & _
                "findStr='" & findStr & "', " & _
                "letterCase=" & letterCase & ", " & _
                "exactMatch=" & exactMatch & ", " & _
                "useRegEx=" & useRegEx
    Debug.Print "  期待結果: " & expectedResult & ", 実際の結果: " & result
    
    If result = expectedResult Then
        Debug.Print "  テスト成功"
        TestCase = True
    Else
        Debug.Print "  テスト失敗"
        TestCase = False
    End If
    
    Debug.Print ""
End Function

