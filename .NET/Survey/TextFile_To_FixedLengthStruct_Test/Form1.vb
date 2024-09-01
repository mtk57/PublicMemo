Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text

Public Class Form1
    <StructLayout(LayoutKind.Sequential)>
    Friend Structure stIN_1
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=5)>
        Public OrderNum As Byte()

        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=4)>
        Public SubNum As Byte()
    End Structure

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Dim stIN_1_size As Integer = Marshal.SizeOf(GetType(stIN_1))
            Dim outputList As New List(Of String)()

            Using sr As New StreamReader("IN_1.txt", Encoding.ASCII)
                Dim line As String = ""
                While (InlineAssignHelper(line, sr.ReadLine())) IsNot Nothing
                    Dim byData As Byte() = Encoding.ASCII.GetBytes(line)
                    Dim pData As IntPtr = Marshal.AllocHGlobal(stIN_1_size)
                    Marshal.Copy(byData, 0, pData, Math.Min(byData.Length, stIN_1_size))
                    Dim stIn1 As stIN_1 = CType(Marshal.PtrToStructure(pData, GetType(stIN_1)), stIN_1)
                    Marshal.FreeHGlobal(pData)

                    ' 構造体の内容を文字列に変換
                    Dim orderNumStr As String = Encoding.ASCII.GetString(stIn1.OrderNum).Trim()
                    Dim subNumStr As String = Encoding.ASCII.GetString(stIn1.SubNum).Trim()
                    outputList.Add($"OrderNum: {orderNumStr}, SubNum: {subNumStr}")
                End While
            End Using

            ' 出力ファイルに書き込み
            File.WriteAllLines("OUT_1.txt", outputList)
        Catch ex As Exception
            MessageBox.Show(ex.ToString())
        End Try

    End Sub

    'VB.NETでは C# の while ((line = reader.ReadLine()) != null) のような記法がないため、メソッドで同等機能を実現
    Private Shared Function InlineAssignHelper(Of T)(ByRef target As T, value As T) As T
        target = value
        Return value
    End Function
End Class
