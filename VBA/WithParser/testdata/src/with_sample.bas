Private Sub Hoge()
    With MyObject 
     .Height = 100 ' Same as MyObject.Height = 100. 
     .Caption = "Hello World" ' Same as MyObject.Caption = "Hello World". 
     With .Font 
      .Color = Red ' Same as MyObject.Font.Color = Red. 
      ret = TestFunc(.Color, .Bold)  'hoge
      ret = TestFunc2( _
               .Color, _
               .Bold)
      ret = TestSub .Color, .Bold   'hoge
      .Bold = True ' Same as MyObject.Font.Bold = True. 
     End With
    End With
End Sub
