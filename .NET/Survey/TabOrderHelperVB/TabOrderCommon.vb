Namespace TabOrderHelper
	''' <summary>
	''' タブオーダー共通クラス
	''' </summary>
	Friend Module Common
		''' <summary>
		''' セパレーター
		''' </summary>
		Public Const SEP As Char = ":"c
	End Module

	''' <summary>
	''' 階層タブインデックスインターフェース
	''' </summary>
	Friend Interface IHasHierarchicalTabIndices
		Inherits System.Windows.Forms.IWin32Window
		Inherits System.Collections.Generic.IEnumerable(Of Integer)
		Inherits System.IComparable
		Inherits System.IComparable(Of IHasHierarchicalTabIndices)
		''' <summary>
		''' 階層タブインデックスをシーケンスで返す
		''' </summary>
		ReadOnly Property HierarchicalTabIndices() As System.Collections.Generic.IEnumerable(Of Integer)
	End Interface

	''' <summary>
	''' ControlNotFoundException
	''' </summary>
	Friend Class ControlNotFoundException
		Inherits System.Exception
		' do nothing
		Private Sub New()
		End Sub

		''' <summary>
		''' コンストラクタ
		''' </summary>
		''' <param name="message">メッセージ</param>
		Public Sub New(message As String)
			' do nothing
			MyBase.New(message)
		End Sub
	End Class

	''' <summary>
	''' Win32APIアクセス用クラス
	''' </summary>
	Friend NotInheritable Class PlatformInvoker
		''' <summary>
		''' GetWindow関数のコマンド
		''' </summary>
		Public Enum GetWindowCmd
			GW_HWNDFIRST = 0
			GW_HWNDLAST = 1
			GW_HWNDNEXT = 2
			GW_HWNDPREV = 3
			GW_OWNER = 4
			GW_CHILD = 5
			GW_ENABLEDPOPUP = 6
		End Enum

		<System.Runtime.InteropServices.DllImport("user32.dll")>
		Public Shared Function GetWindow(ByVal hwd As System.IntPtr, ByVal uCmd As UInteger) As System.IntPtr
		End Function
	End Class


	''' <summary>
	''' ソート順
	''' </summary>
	Friend Enum Sort
		Asc
		Desc
	End Enum

	''' <summary>
	''' 階層タブインデックスのソートヘルパークラス
	''' </summary>
	Friend Class SortHelperOfHierarchicalTabIndices
		Implements System.Collections.Generic.IComparer(Of IHasHierarchicalTabIndices)

		Private _toggle As Integer = 1

		Public Sub New()
			' do nothing
		End Sub

		Public Sub New(sort As Sort)
			Select Case sort
				Case Sort.Asc
			' do nothing
				Case Sort.Desc
					_toggle = -1
				Case Else
					_toggle = 1
			End Select
		End Sub

		Public Function Compare(ByVal x As IHasHierarchicalTabIndices, ByVal y As IHasHierarchicalTabIndices) As Integer Implements IComparer(Of IHasHierarchicalTabIndices).Compare
			Using enumerator1 = x.GetEnumerator()
				Using enumerator2 = y.GetEnumerator()
					Dim e1 = enumerator1.MoveNext()
					Dim e2 = enumerator2.MoveNext()

					While e1 AndAlso e2
						Dim result = enumerator1.Current.CompareTo(enumerator2.Current) * _toggle
						If result <> 0 Then
							Return result
						End If

						e1 = enumerator1.MoveNext()
						e2 = enumerator2.MoveNext()
					End While

					' 比較対象が無くなった

					If Not e1 AndAlso Not e2 Then
						' TabIndexの階層構造に全く同じ値が設定されていた場合はZオーダーで比較する
						Return CompareZOrder(x.Handle, y.Handle)
					End If

					If Not e1 Then
						Return -1 * _toggle
					End If

					If Not e2 Then
						Return 1 * _toggle
					End If
				End Using
			End Using
			Return 0
		End Function


		Private Function CompareZOrder(ByVal hwdx As System.IntPtr, ByVal hwdy As System.IntPtr) As Integer
			Dim h = PlatformInvoker.GetWindow(hwdx, CUInt(PlatformInvoker.GetWindowCmd.GW_HWNDNEXT))
			While h <> System.IntPtr.Zero
				If h = hwdy Then
					Return -1 * _toggle
				End If

				h = PlatformInvoker.GetWindow(h, CUInt(PlatformInvoker.GetWindowCmd.GW_HWNDNEXT))
			End While

			h = PlatformInvoker.GetWindow(hwdx, CUInt(PlatformInvoker.GetWindowCmd.GW_HWNDPREV))
			While h <> System.IntPtr.Zero
				If h = hwdy Then
					Return 1 * _toggle
				End If

				h = PlatformInvoker.GetWindow(h, CUInt(PlatformInvoker.GetWindowCmd.GW_HWNDPREV))
			End While

			Return 0
		End Function

	End Class

End Namespace