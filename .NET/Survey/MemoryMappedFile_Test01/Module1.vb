Imports System.IO
Imports System.IO.MemoryMappedFiles
Imports System.Runtime.InteropServices
Imports System.Text

'https://potisan-programming-memo.hatenablog.jp/entry/2014/05/21/000000
'https://learn.microsoft.com/ja-jp/dotnet/api/system.io.memorymappedfiles.memorymappedfile?view=net-8.0

Module BinaryToStructExample
    Const FILE_PATH As String = "C:\_tmp\aa\ConsoleApp1\testdata\test.txt"

    '<StructLayout(LayoutKind.Sequential, Pack:=1)>
    'Public Structure DataStruct
    '    <MarshalAs(UnmanagedType.ByValArray, SizeConst:=10)>
    '    Public Field As Byte()

    '    <MarshalAs(UnmanagedType.ByValArray, SizeConst:=2)>
    '    Public CrLf As Byte()

    '    Public Sub New(dummy As Integer)
    '        Field = New Byte(9) {}
    '        CrLf = New Byte(1) {}
    '    End Sub


    'End Structure

    Private Function ConvertByteArrayToString(bytes As Byte()) As String
        Dim result As New Text.StringBuilder(bytes.Length)
        For Each b In bytes
            If b = 0 Then
                result.Append(ChrW(0))  ' ヌル文字を明示的に追加
            Else
                result.Append(ChrW(b))
            End If
        Next
        Return result.ToString()
    End Function





    '----------------------------------------------
    'Sample 03
    Const MAP_NAME As String = "CustomStructMapping"
    Const ROW_SIZE As Integer = 12 ' 10 bytes data + 2 bytes CRLF
    Const TOTAL_ROWS As Integer = 4

    <StructLayout(LayoutKind.Sequential, Pack:=1)>
    Public Structure CustomStruct
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=5)>
        Public Field1 As Byte()

        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=5)>
        Public Field2 As Byte()

        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=2)>
        Public Field3 As Byte()
    End Structure

    Sub Main()
        Using mmf As MemoryMappedFile = MemoryMappedFile.CreateFromFile(FILE_PATH, FileMode.Open, MAP_NAME)
            Using accessor As MemoryMappedViewAccessor = mmf.CreateViewAccessor()
                Dim buffer(ROW_SIZE - 1) As Byte
                For row As Integer = 0 To TOTAL_ROWS - 1
                    accessor.ReadArray(row * ROW_SIZE, buffer, 0, ROW_SIZE)
                    Dim customStruct As CustomStruct = ByteArrayToStruct(buffer)
                    PrintStructContent(customStruct, row + 1)
                Next
            End Using
        End Using
        Console.WriteLine("Processing complete")
    End Sub

    Private Function ByteArrayToStruct(bytes As Byte()) As CustomStruct
        Dim size As Integer = Marshal.SizeOf(GetType(CustomStruct))
        Dim ptr As IntPtr = Marshal.AllocHGlobal(size)
        Try
            Marshal.Copy(bytes, 0, ptr, size)
            Return CType(Marshal.PtrToStructure(ptr, GetType(CustomStruct)), CustomStruct)
        Finally
            Marshal.FreeHGlobal(ptr)
        End Try
    End Function

    Private Sub PrintStructContent(customStruct As CustomStruct, rowNumber As Integer)
        Console.WriteLine($"Row {rowNumber}:")
        Console.WriteLine($"  Field1: {BitConverter.ToString(customStruct.Field1)}")
        Console.WriteLine($"  Field2: {BitConverter.ToString(customStruct.Field2)}")
        Console.WriteLine($"  Field3: {BitConverter.ToString(customStruct.Field3)}")
        Console.WriteLine()
    End Sub



    '----------------------------------------------
    'Sample 02

    'Const MAP_NAME As String = "LargeFileMapping"

    'Sub Main()
    '    Dim fileInfo As New FileInfo(FILE_PATH)
    '    Dim fileSize As Long = fileInfo.Length
    '    Dim structSize As Integer = Marshal.SizeOf(GetType(DataStruct))

    '    Using mmf As MemoryMappedFile = MemoryMappedFile.CreateFromFile(FILE_PATH, FileMode.Open, MAP_NAME, fileSize)
    '        Using accessor As MemoryMappedViewAccessor = mmf.CreateViewAccessor()
    '            Dim position As Long = 0
    '            While position < fileSize
    '                Dim dataStruct As New DataStruct()
    '                'dataStruct.Field = New Byte(9) {}  ' 10バイトの配列を初期化
    '                dataStruct.Field = New Byte(9) {}  ' 10バイトの配列を初期化
    '                dataStruct.CrLf = New Byte(1) {}  ' 10バイトの配列を初期化
    '                accessor.ReadArray(position, dataStruct.Field, 0, structSize)
    '                ProcessStruct(dataStruct)
    '                position += 10  ' 10バイト進める
    '            End While
    '        End Using
    '    End Using
    '    Console.WriteLine("Processing complete")
    'End Sub

    'Sub ProcessStruct(dataStruct As DataStruct)
    '    Console.WriteLine(ConvertByteArrayToString(dataStruct.Field))

    '    ' ここで個々の構造体を処理
    '    ' 例: データベースに保存、集計、変換など
    'End Sub




    '----------------------------------------------
    'Sample 01

    'Structure st2
    '    <VBFixedString(10)> Dim Field3 As String
    'End Structure

    'Sub Main()
    '    Dim st22 As st2 = New st2
    '    Dim fnum As Integer = FreeFile()
    '    FileOpen(fnum, PATH, OpenMode.Binary, OpenAccess.Read)
    '    FileGet(fnum, st22)
    '    FileClose(fnum)


    '    Dim fileInfo As New FileInfo(PATH)
    '    Dim structSize As Integer = Marshal.SizeOf(GetType(MyStruct))

    '    Using fileStream As New FileStream(PATH, FileMode.Open, FileAccess.Read)
    '        Dim buffer(structSize - 1) As Byte
    '        fileStream.Read(buffer, 0, structSize)
    '        Dim handle As GCHandle = GCHandle.Alloc(buffer, GCHandleType.Pinned)
    '        Dim myStruct As MyStruct
    '        Try
    '            myStruct = CType(Marshal.PtrToStructure(handle.AddrOfPinnedObject(), GetType(MyStruct)), MyStruct)
    '            ' バイト配列を文字列に変換
    '            'Dim field3String As String = Encoding.UTF8.GetString(myStruct.Field3).TrimEnd(CChar(vbNullChar))
    '            Dim field3String As String = ConvertByteArrayToString(myStruct.Field3)

    '            Console.WriteLine($"Field3: '{field3String}'")
    '            Console.WriteLine($"Field3 (Hex): '{BitConverter.ToString(myStruct.Field3)}'")
    '        Finally
    '            handle.Free()
    '        End Try
    '    End Using
    'End Sub


End Module