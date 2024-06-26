Imports System.Threading
Imports System.IO
Imports System
Imports System.Reflection

Namespace Logger

    Public Class Logger
        Public Shared Sub WriteLog(ByVal contents As String)
            Dim file_num As Integer
            file_num = FreeFile()
            FileOpen(file_num, "Logger.log", OpenMode.Append)
            PrintLine(file_num, Format(Today, "yyyy/mm/dd") & " " & Format(Now, "hh:mm:ss") & ":" & contents)
            FileClose(file_num)
            file_num = -1
        End Sub

    End Class
End Namespace