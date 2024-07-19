Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            ' 文字列配列を作成し、CopyArrayEx関数をテスト
            Dim a(1) As String
            a(0) = "aa"
            a(1) = "bb"
            Dim b() As String
            b = CopyArrayEx(a)

            ' すべてのテストを実行
            TestMain()
        Catch ex As Exception
            ' エラーが発生した場合、メッセージボックスで表示
            MessageBox.Show(ex.ToString())
        End Try
    End Sub

    ' ジェネリック型の配列をコピーする関数
    Public Function CopyArrayEx(Of T)(src As T()) As T()
        ' 入力配列がNullの場合、Nullを返す
        If src Is Nothing Then
            Return Nothing
        End If
        ' 新しい配列を作成し、元の配列の内容をコピー
        Dim dest(src.Length - 1) As T
        Array.Copy(src, dest, src.Length)
        Return dest
    End Function

    ' 多次元配列やジャグ配列を含む任意の配列をコピーする関数
    Public Function CopyArrayEx(Of T)(src As Array) As Array
        ' 入力配列がNullの場合、Nullを返す
        If src Is Nothing Then
            Return Nothing
        End If
        ' 配列のランクと各次元の長さ、下限を取得
        Dim rank As Integer = src.Rank
        Dim lengths(rank - 1) As Integer
        Dim lowerBounds(rank - 1) As Integer
        For i As Integer = 0 To rank - 1
            lengths(i) = src.GetLength(i)
            lowerBounds(i) = src.GetLowerBound(i)
        Next
        ' 新しい配列を作成し、元の配列の内容をコピー
        Dim dest As Array = Array.CreateInstance(GetType(T), lengths, lowerBounds)
        Array.Copy(src, dest, src.Length)
        Return dest
    End Function

    ' すべてのテストケースを実行する関数
    Private Sub TestMain()
        TestCopyArrayEx_Null()
        TestCopyArrayEx_EmptyArray()
        TestCopyArrayEx_SingleDimensionIntArray()
        TestCopyArrayEx_MultiDimensionIntArray()
        TestCopyArrayEx_CustomLowerBound()
        TestCopyArrayEx_StringArray()
        TestCopyArrayEx_DoubleArray()
        TestCopyArrayEx_BooleanArray()
        TestCopyArrayEx_DateTimeArray()
        TestCopyArrayEx_CustomStructArray()
        TestCopyArrayEx_JaggedArray()
        TestCopyArrayEx_LargeArray()

        MessageBox.Show("All tests completed.")
    End Sub

    ' Null配列のテスト
    Sub TestCopyArrayEx_Null()
        Dim result As Array = CopyArrayEx(Of Integer)(Nothing)
        AssertIsNull(result, "TestCopyArrayEx_Null")
    End Sub

    ' 空の配列のテスト
    Sub TestCopyArrayEx_EmptyArray()
        Dim source As Integer() = {}
        Dim result As Array = CopyArrayEx(Of Integer)(source)
        AssertNotNull(result, "TestCopyArrayEx_EmptyArray")
        AssertEqual(0, result.Length, "TestCopyArrayEx_EmptyArray: Length")
    End Sub

    ' 1次元整数配列のテスト
    Sub TestCopyArrayEx_SingleDimensionIntArray()
        Dim source As Integer() = {1, 2, 3, 4, 5}
        Dim result As Array = CopyArrayEx(Of Integer)(source)
        AssertNotNull(result, "TestCopyArrayEx_SingleDimensionIntArray")
        AssertEqual(source.Length, result.Length, "TestCopyArrayEx_SingleDimensionIntArray: Length")
        AssertArrayEqual(source, DirectCast(result, Integer()), "TestCopyArrayEx_SingleDimensionIntArray: Content")
    End Sub

    ' 多次元整数配列のテスト
    Sub TestCopyArrayEx_MultiDimensionIntArray()
        Dim source(,) As Integer = {{1, 2}, {3, 4}, {5, 6}}
        Dim result As Array = CopyArrayEx(Of Integer)(source)
        AssertNotNull(result, "TestCopyArrayEx_MultiDimensionIntArray")
        AssertEqual(source.GetLength(0), result.GetLength(0), "TestCopyArrayEx_MultiDimensionIntArray: Dimension 0")
        AssertEqual(source.GetLength(1), result.GetLength(1), "TestCopyArrayEx_MultiDimensionIntArray: Dimension 1")
        For i As Integer = 0 To source.GetLength(0) - 1
            For j As Integer = 0 To source.GetLength(1) - 1
                AssertEqual(source(i, j), result.GetValue(i, j), $"TestCopyArrayEx_MultiDimensionIntArray: Element ({i},{j})")
            Next
        Next
    End Sub

    ' カスタム下限を持つ配列のテスト
    Sub TestCopyArrayEx_CustomLowerBound()
        Dim source As Array = Array.CreateInstance(GetType(Integer), New Integer() {3}, New Integer() {2})
        source.SetValue(10, 2)
        source.SetValue(20, 3)
        source.SetValue(30, 4)

        Dim result As Array = CopyArrayEx(Of Integer)(source)
        AssertNotNull(result, "TestCopyArrayEx_CustomLowerBound")
        AssertEqual(source.Length, result.Length, "TestCopyArrayEx_CustomLowerBound: Length")
        AssertEqual(source.GetLowerBound(0), result.GetLowerBound(0), "TestCopyArrayEx_CustomLowerBound: LowerBound")
        For i As Integer = source.GetLowerBound(0) To source.GetUpperBound(0)
            AssertEqual(source.GetValue(i), result.GetValue(i), $"TestCopyArrayEx_CustomLowerBound: Element {i}")
        Next
    End Sub

    ' 文字列配列のテスト
    Sub TestCopyArrayEx_StringArray()
        Dim source As String() = {"Hello", "World", "Test"}
        Dim result As Array = CopyArrayEx(Of String)(source)
        AssertNotNull(result, "TestCopyArrayEx_StringArray")
        AssertEqual(source.Length, result.Length, "TestCopyArrayEx_StringArray: Length")
        AssertArrayEqual(source, DirectCast(result, String()), "TestCopyArrayEx_StringArray: Content")
    End Sub

    ' Double配列のテスト
    Sub TestCopyArrayEx_DoubleArray()
        Dim source As Double() = {1.1, 2.2, 3.3, 4.4, 5.5}
        Dim result As Array = CopyArrayEx(Of Double)(source)
        AssertNotNull(result, "TestCopyArrayEx_DoubleArray")
        AssertEqual(source.Length, result.Length, "TestCopyArrayEx_DoubleArray: Length")
        AssertArrayEqual(source, DirectCast(result, Double()), "TestCopyArrayEx_DoubleArray: Content")
    End Sub

    ' Boolean配列のテスト
    Sub TestCopyArrayEx_BooleanArray()
        Dim source As Boolean() = {True, False, True, True, False}
        Dim result As Array = CopyArrayEx(Of Boolean)(source)
        AssertNotNull(result, "TestCopyArrayEx_BooleanArray")
        AssertEqual(source.Length, result.Length, "TestCopyArrayEx_BooleanArray: Length")
        AssertArrayEqual(source, DirectCast(result, Boolean()), "TestCopyArrayEx_BooleanArray: Content")
    End Sub

    ' DateTime配列のテスト
    Sub TestCopyArrayEx_DateTimeArray()
        Dim source As DateTime() = {DateTime.Now, DateTime.UtcNow, New DateTime(2000, 1, 1)}
        Dim result As Array = CopyArrayEx(Of DateTime)(source)
        AssertNotNull(result, "TestCopyArrayEx_DateTimeArray")
        AssertEqual(source.Length, result.Length, "TestCopyArrayEx_DateTimeArray: Length")
        AssertArrayEqual(source, DirectCast(result, DateTime()), "TestCopyArrayEx_DateTimeArray: Content")
    End Sub

    ' カスタム構造体配列のテスト
    Sub TestCopyArrayEx_CustomStructArray()
        Dim source As MyStruct() = {New MyStruct(1, "A"), New MyStruct(2, "B"), New MyStruct(3, "C")}
        Dim result As Array = CopyArrayEx(Of MyStruct)(source)
        AssertNotNull(result, "TestCopyArrayEx_CustomStructArray")
        AssertEqual(source.Length, result.Length, "TestCopyArrayEx_CustomStructArray: Length")
        Dim resultArray As MyStruct() = DirectCast(result, MyStruct())
        For i As Integer = 0 To source.Length - 1
            AssertEqual(source(i).Id, resultArray(i).Id, $"TestCopyArrayEx_CustomStructArray: Element {i} Id")
            AssertEqual(source(i).Name, resultArray(i).Name, $"TestCopyArrayEx_CustomStructArray: Element {i} Name")
        Next
    End Sub

    ' ジャグ配列のテスト
    Sub TestCopyArrayEx_JaggedArray()
        Dim source As Integer()() = {New Integer() {1, 2, 3}, New Integer() {4, 5}, New Integer() {6, 7, 8, 9}}
        Dim result As Array = CopyArrayEx(Of Integer())(source)
        AssertNotNull(result, "TestCopyArrayEx_JaggedArray")
        AssertEqual(source.Length, result.Length, "TestCopyArrayEx_JaggedArray: Length")
        Dim resultArray As Integer()() = DirectCast(result, Integer()())
        For i As Integer = 0 To source.Length - 1
            AssertNotNull(resultArray(i), $"TestCopyArrayEx_JaggedArray: SubArray {i}")
            AssertEqual(source(i).Length, resultArray(i).Length, $"TestCopyArrayEx_JaggedArray: SubArray {i} Length")
            For j As Integer = 0 To source(i).Length - 1
                AssertEqual(source(i)(j), resultArray(i)(j), $"TestCopyArrayEx_JaggedArray: Element ({i},{j})")
            Next
        Next
    End Sub

    ' 大規模配列のコピーとパフォーマンステスト
    Sub TestCopyArrayEx_LargeArray()
        ' 大きな配列を生成（1000万要素）
        Const ArraySize As Integer = 10000000
        Dim source(ArraySize - 1) As Integer
        For i As Integer = 0 To ArraySize - 1
            source(i) = i
        Next

        ' CopyArrayEx の実行時間を測定
        Dim swCopyArrayEx As New Stopwatch()
        swCopyArrayEx.Start()
        Dim resultCopyArrayEx As Array = CopyArrayEx(Of Integer)(source)
        swCopyArrayEx.Stop()

        ' 標準的な配列コピー方法（Array.Copy）の実行時間を測定
        Dim swArrayCopy As New Stopwatch()
        Dim resultArrayCopy(ArraySize - 1) As Integer
        swArrayCopy.Start()
        Array.Copy(source, resultArrayCopy, ArraySize)
        swArrayCopy.Stop()

        ' 結果を検証
        AssertNotNull(resultCopyArrayEx, "TestCopyArrayEx_LargeArray: CopyArrayEx result")
        AssertEqual(ArraySize, resultCopyArrayEx.Length, "TestCopyArrayEx_LargeArray: CopyArrayEx length")
        AssertEqual(source(0), DirectCast(resultCopyArrayEx, Integer())(0), "TestCopyArrayEx_LargeArray: First element")
        AssertEqual(source(ArraySize - 1), DirectCast(resultCopyArrayEx, Integer())(ArraySize - 1), "TestCopyArrayEx_LargeArray: Last element")

        ' 実行時間を比較して出力
        Console.WriteLine($"CopyArrayEx execution time: {swCopyArrayEx.ElapsedMilliseconds} ms")
        Console.WriteLine($"Array.Copy execution time: {swArrayCopy.ElapsedMilliseconds} ms")
        Console.WriteLine($"Time difference: {swCopyArrayEx.ElapsedMilliseconds - swArrayCopy.ElapsedMilliseconds} ms")

        ' パフォーマンスの許容範囲を定義（例: CopyArrayExがArray.Copyの1.5倍以内であること）
        Const PerformanceThreshold As Double = 1.5
        Dim performanceRatio As Double = swCopyArrayEx.ElapsedMilliseconds / swArrayCopy.ElapsedMilliseconds
        AssertLessThan(performanceRatio, PerformanceThreshold, $"TestCopyArrayEx_LargeArray: Performance ratio ({performanceRatio:F2}) should be less than {PerformanceThreshold}")

        Console.WriteLine("TestCopyArrayEx_LargeArray: Passed")
    End Sub

    ' 値がNullであることを確認するアサーション
    Private Sub AssertIsNull(Of T)(value As T, message As String)
        If value IsNot Nothing Then
            Throw New Exception($"{message}: Expected null, but got non-null value")
        End If
        Console.WriteLine($"{message}: Passed")
    End Sub

    ' 値がNullでないことを確認するアサーション
    Private Sub AssertNotNull(Of T)(value As T, message As String)
        If value Is Nothing Then
            Throw New Exception($"{message}: Expected non-null value, but got null")
        End If
        Console.WriteLine($"{message}: Passed")
    End Sub

    ' 二つの値が等しいことを確認するアサーション
    Private Sub AssertEqual(Of T)(expected As T, actual As T, message As String)
        If Not EqualityComparer(Of T).Default.Equals(expected, actual) Then
            Throw New Exception($"{message}: Expected {expected}, but got {actual}")
        End If
        Console.WriteLine($"{message}: Passed")
    End Sub

    ' 二つの配列が等しいことを確認するアサーション
    Private Sub AssertArrayEqual(Of T)(expected As T(), actual As T(), message As String)
        If expected.Length <> actual.Length Then
            Throw New Exception($"{message}: Array lengths do not match. Expected {expected.Length}, but got {actual.Length}")
        End If
        For i As Integer = 0 To expected.Length - 1
            If Not EqualityComparer(Of T).Default.Equals(expected(i), actual(i)) Then
                Throw New Exception($"{message}: Element at index {i} does not match. Expected {expected(i)}, but got {actual(i)}")
            End If
        Next
        Console.WriteLine($"{message}: Passed")
    End Sub

    ' 値が指定された値より小さいことを確認するアサーション
    Private Sub AssertLessThan(actual As Double, expected As Double, message As String)
        If actual >= expected Then
            Throw New Exception($"{message}: Expected less than {expected}, but got {actual}")
        End If
        Console.WriteLine($"{message}: Passed")
    End Sub

End Class

' カスタム構造体の定義
Public Structure MyStruct
    Public Id As Integer
    Public Name As String

    Public Sub New(id As Integer, name As String)
        Me.Id = id
        Me.Name = name
    End Sub
End Structure