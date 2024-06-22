Module SimpleLogger

    Public Sub WriteLog(ByVal contents As String)
        Dim file_num As Integer
        file_num = FreeFile()
        FileOpen(file_num, "SimpleLogger.log", OpenMode.Append)
        PrintLine(file_num, VB6.Format(Today, "yyyy/mm/dd") & " " & VB6.Format(Now, "hh:mm:ss") & ":" & contents)
        FileClose(file_num)
        file_num = -1
    End Sub
End Module
