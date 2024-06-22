Module SimpleLogger

    Public Sub WriteLog(ByVal contents As String)
        Dim file_num As Integer
        file_num = FreeFile()
        FileOpen(file_num, "SimpleLogger.log", OpenMode.Append)
        PrintLine(file_num, Format(Today, "yyyy/mm/dd") & " " & Format(Now, "hh:mm:ss") & ":" & contents)
        FileClose(file_num)
        file_num = -1
    End Sub
End Module
