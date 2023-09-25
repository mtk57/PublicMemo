Namespace TabOrderHelper
	''' <summary>
	''' タブオーダーモデル
	''' </summary>
	Friend NotInheritable Class TabOrderModel
		Implements IHasHierarchicalTabIndices

		Private _hierarchicalTabIndices As System.Collections.Generic.IEnumerable(Of Integer)

		''' <summary>
		''' 前のコントロールモデル
		''' </summary>
		Public Property PrevControl() As TabOrderModel

		''' <summary>
		''' カレントコントロール
		''' </summary>
		Public ReadOnly Property Control() As System.Windows.Forms.Control

		''' <summary>
		''' 次のコントロールモデル
		''' </summary>
		Public Property NextControl() As TabOrderModel

		''' <summary>
		''' タブインデックス文字列
		''' 階層表記はで親子をデリミタで区切る
		''' </summary>
		Public ReadOnly Property IndexString() As String

		''' <summary>
		''' 最後の階層の親のタブインデックス
		''' </summary>
		Public ReadOnly Property ParentLastIndex() As Integer

		''' <summary>
		''' 最後の階層のタブインデックス
		''' 重複の可能性あり
		''' 重複している場合、Zオーダーで順序を決定する
		''' </summary>
		Public ReadOnly Property LastIndex() As Integer

		''' <summary>
		''' 内部的にナンバリングした重複無しのタブインデックス
		''' </summary>
		Public Property UniqueTabIndex() As Integer?

		''' <summary>
		''' コンテナ系コントロールか否か
		''' </summary>
		Public ReadOnly Property IsContainer() As Boolean

		''' <summary>
		''' ラジオボタンコントロールか否か
		''' </summary>
		Public ReadOnly Property IsRadioButton() As Boolean

		''' <summary>
		''' ユーザーコントロールの子供か否か
		''' </summary>
		Public ReadOnly Property IsUserControlChild() As Boolean
			Get
				Return TypeOf Control.Parent Is System.Windows.Forms.UserControl
			End Get
		End Property

		''' <summary>
		''' コントロールのウィンドウハンドル
		''' Zオーダー判定時に必要
		''' </summary>
		Public ReadOnly Property Handle() As System.IntPtr Implements IWin32Window.Handle
			Get
				Return Control.Handle
			End Get
		End Property

		Private Sub New()
			' do nothing
		End Sub

		''' <summary>
		''' コンストラクタ
		''' </summary>
		''' <param name="c">コントロール</param>
		Public Sub New(c As System.Windows.Forms.Control)
			' 階層タブインデックスを取得する
			_hierarchicalTabIndices = GetHierarchicalTabindices(c)

			PrevControl = Nothing
			Control = c
			NextControl = Nothing
			IndexString = GetHierarchicalTabIndicesString(Control)
			ParentLastIndex = GetPreviousNumber(IndexString)
			LastIndex = GetLastNumber(IndexString)
			UniqueTabIndex = Nothing
			IsRadioButton = TypeOf Control Is System.Windows.Forms.RadioButton
		End Sub

		Public Overrides Function ToString() As String
			If PrevControl Is Nothing Then
				Return $"Name={Control.Name}" & vbTab &
					   $"PrevUniqueTabIndex=" & vbTab &
					   $"TabIndex={Control.TabIndex}" & vbTab &
					   $"NextUniqueTabIndex=" & vbTab &
					   $"IndexString={IndexString}" & vbTab &
					   $"ParentLastIndex={ParentLastIndex}" & vbTab &
					   $"LastIndex={LastIndex}" & vbTab &
					   $"UniqueTabIndex={UniqueTabIndex}" & vbTab &
					   $"IsContainer={IsContainer}" & vbTab &
					   $"IsRadioButton={IsRadioButton}" & vbTab &
					   $"IsUserControlChild={IsUserControlChild}"
			End If

			Return $"Name={Control.Name}" & vbTab &
				   $"PrevUniqueTabIndex={PrevControl.UniqueTabIndex}" & vbTab &
				   $"TabIndex={Control.TabIndex}" & vbTab &
				   $"NextUniqueTabIndex={NextControl.UniqueTabIndex}" & vbTab &
				   $"IndexString={IndexString}" & vbTab &
				   $"ParentLastIndex={ParentLastIndex}" & vbTab &
				   $"LastIndex={LastIndex}" & vbTab &
				   $"UniqueTabIndex={UniqueTabIndex}" & vbTab &
				   $"IsContainer={IsContainer}" & vbTab &
				   $"IsRadioButton={IsRadioButton}" & vbTab &
				   $"IsUserControlChild={IsUserControlChild}"
		End Function

		Public ReadOnly Property HierarchicalTabIndices() As System.Collections.Generic.IEnumerable(Of Integer) Implements IHasHierarchicalTabIndices.HierarchicalTabIndices
			Get
				Return _hierarchicalTabIndices
			End Get
		End Property

		Public Function CompareTo(obj As Object) As Integer Implements IComparable.CompareTo
			Return CompareTo(CType(obj, IHasHierarchicalTabIndices))
		End Function

		Public Function CompareTo(other As IHasHierarchicalTabIndices) As Integer Implements IComparable(Of IHasHierarchicalTabIndices).CompareTo
			Return New SortHelperOfHierarchicalTabIndices().Compare(Me, other)
		End Function

		Public Function GetEnumerator() As IEnumerator(Of Integer) Implements IEnumerable(Of Integer).GetEnumerator
			Return Me.HierarchicalTabIndices.GetEnumerator()
		End Function

		Function GetEnumerator_() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
			Return CType(Me.HierarchicalTabIndices.GetEnumerator(), System.Collections.IEnumerator)
		End Function

		''' <summary>
		''' 階層構造を持ったコントロールのタブインデックスシーケンスを返す
		''' </summary>
		''' <param name="control">コントロール</param>
		''' <returns>タブインデックスシーケンス</returns>
		Private Iterator Function GetHierarchicalTabindices(control As System.Windows.Forms.Control) As System.Collections.Generic.IEnumerable(Of Integer)
			Dim s As New System.Collections.Generic.Stack(Of Integer)()
			s.Push(control.TabIndex)
			Dim parent As System.Windows.Forms.Control = control.Parent
			While IsParent(parent)
				s.Push(parent.TabIndex)
				parent = parent.Parent
			End While

			While s.Count <> 0
				Yield s.Pop()
			End While
		End Function

		''' <summary>
		''' 階層構造を持ったコントロールのタブインデックスを文字列で返す
		''' 例.
		'''   Form
		'''     GroupBox     0
		'''        Button1   1
		'''        TextBox1  2
		'''     Button2      3
		'''     
		'''    はそれぞれ以下が返る
		'''    Button1="0:1"
		'''    TextBox1="0:2"
		'''    Button2="3"
		''' </summary>
		''' <param name="control">コントロール</param>
		''' <returns>タブインデックス</returns>
		Private Function GetHierarchicalTabIndicesString(control As System.Windows.Forms.Control) As String
			Dim sb As New System.Text.StringBuilder()
			For Each item As Integer In GetHierarchicalTabindices(control)
				sb.AppendFormat("{0}" & Common.SEP, item.ToString())
			Next
			Return System.Text.RegularExpressions.Regex.Replace(sb.ToString(), Common.SEP & "$", "")
		End Function

		''' <summary>
		''' 対象コントロールが親コントロールか否か
		''' </summary>
		''' <param name="target">対象コントロール</param>
		''' <returns>True:親コントロール, False:親コントロールではない</returns>
		Private Function IsParent(target As System.Windows.Forms.Control) As Boolean
			If target Is Nothing Then
				Return False
			End If
			If TypeOf target Is System.Windows.Forms.Form Then
				Return False
			End If
			Return True
		End Function

		''' <summary>
		''' タブインデックス文字列の最後の階層の1つ上を返す
		''' 例1:"1:2:3"の場合2が返る
		''' 例2:"3"の場合-1が返る
		''' </summary>
		''' <param name="indexString">タブインデックス文字列</param>
		''' <returns>最後の階層の1つ上の値</returns>
		Private Function GetPreviousNumber(indexString As String) As Integer
			Dim numbers() As String = indexString.Split(Common.SEP)
			Dim length As Integer = numbers.Length
			Dim secondLastNumber As Integer = -1
			'コンテナに内包されていない場合
			If length > 1 Then
				Integer.TryParse(numbers(length - 2), secondLastNumber)
			End If
			Return secondLastNumber
		End Function

		''' <summary>
		''' タブインデックス文字列の最後の階層を返す
		''' 例1:"1:2:3"の場合3が返る
		''' </summary>
		''' <param name="indexString">タブインデックス文字列</param>
		''' <returns>最後の階層の値</returns>
		Private Function GetLastNumber(indexString As String) As Integer
			Dim parts() As String = indexString.Split(Common.SEP)
			Dim lastPart As String = parts(parts.Length - 1)
			Return Integer.Parse(lastPart)
		End Function
	End Class
End Namespace