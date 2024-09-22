Option Explicit

●関数定義より上

Public Function Func1(ByVal a1 As String) As Long

●Function1 引数あり・戻り値あり

End Function

Public Function Func2(ByVal a1 As String)

●Function2 引数あり・戻り値なし

End Function

Public Function Func3()

●Function3 引数なし・戻り値なし

End Function

Public Function Func4() As Long

●Function4 引数なし・戻り値あり

End Function

Public Function Func5(ByVal a As Long, ByRef b As Long)

●Function5 引数あり(複数・改行なし)・戻り値なし

End Function

Public Function Func6(ByVal a As Long, _
                      ByRef b As Long)

●Function6 引数あり(複数・改行あり)・戻り値なし

End Function

Public Function Func7(ByVal a As Long, _
                      ByRef b As Long) As Long

●Function7 引数あり(複数・改行あり)・戻り値あり

End Function

Public Function Func8(ByVal a As Long, _
                      ByRef b As Long _
                      ) As Long

●Function8 引数あり(複数・改行あり)・戻り値あり2

End Function

Public Function Func9(ByVal a As Long, _
                      ByRef b As Long _
                      ) As Long()

●Function9 引数あり(複数・改行あり)・戻り値あり3

End Function

Public Function Func10(ByVal a() As Long, _
                      ByRef b As Long() _
                      ) As Long()

●Function10 引数あり(複数・改行あり)・戻り値あり10

End Function

Public Sub Sub1(ByVal a1 As String)

●Sub1 引数あり

End Sub

Public Sub Sub2()

●Sub2 引数なし

End Sub

Public Sub Sub3(ByVal a As Long, ByRef b As Long)

●Sub3 引数あり(複数・改行なし)

End Sub

Public Sub Sub4(ByVal a As Long, _
                ByRef b As Long)

●Sub4 引数あり(複数・改行あり)

End Sub

Public Sub Sub5(ByVal a() As Long, _
                ByRef b As Long())

●Sub5 引数あり(複数・改行あり)2

End Sub

Public Sub 日本語関数1()

●日本語関数1

End Sub

Public Sub Method日本語関数2()

●Func日本語関数2

End Sub

Public Function 日本語関数3()

●日本語関数3

End Function

Public Function Method日本語関数4()

●Func日本語関数4

End Function

●関数定義より下
