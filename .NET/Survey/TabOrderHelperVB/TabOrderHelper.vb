Imports System.Linq
'
' デザイナーの「表示/タブオーダー」のように、階層化されたタブインデックスをリストで管理する
'
' 参考:https://zecl.hatenablog.com/entry/20090226/p1
'      https://atmarkit.itmedia.co.jp/fdotnet/dotnettips/243winkeyproc/winkeyproc.html
'

Namespace TabOrderHelper
	''' <summary>
	''' タブオーダーヘルパークラス
	''' 
	''' [本クラスを使用する場合の注意点]
	''' 1.コンテナ系コントロールは以下のみ対応する。
	'''   Panel
	'''   GroupBox
	'''
	''' [使用例]
	''' Public Class Form1
	'''     Private _helper As TabOrderHelper.TabOrderHelper = Nothing
	'''
	'''     Private Sub Form1_Load(sender As Object, e As EventArgs)
	''' 	    _helper = New TabOrderHelper.TabOrderHelper(Me)
	''' 	End Sub
	''' 
	'''     Protected Overrides Function ProcessCmdKey(ByRef msg As Message, ByVal keyData As Keys) As Boolean
	'''         Dim control = Me.ActiveControl
	''' 
	'''         If keyData = Keys.Tab Then
	''' 	        ' TABキーが押されたときの処理
	''' 
	''' 	        Dim nextControl = _helper.GetNextControl(control)
	'''			    nextControl.Focus()
	'''			    Return True ' イベントを処理済みとしてマークする
	'''
	'''		    ElseIf keyData = (Keys.Shift Or Keys.Tab) Then
	'''			    ' SHIFT+TABキーが押されたときの処理
	'''
	'''			    Dim prevControl = _helper.GetNextControl(control, False)
	'''			    prevControl.Focus()
	'''			    Return True ' イベントを処理済みとしてマークする
	'''
	'''		    End If
	'''
	'''		    Return MyBase.ProcessCmdKey(msg, keyData)
	'''	    End Function
	'''
	''' End Class
	''' 
	''' </summary>
	Public NotInheritable Class TabOrderHelper
		Private _modelList As System.Collections.Generic.List(Of TabOrderModel)

		Private Sub New()
			' do nothing
		End Sub

		''' <summary>
		''' コンストラクタ
		''' </summary>
		''' <param name="form">フォーム</param>
		Public Sub New(
			form As System.Windows.Forms.Control)

			Update(form)

		End Sub

		''' <summary>
		''' カレントコントロールの次(もしくは前)のコントロールを返す
		''' </summary>
		''' <param name="control">カレントコントロール</param>
		''' <param name="forward">検索方向  True:次のコントロール、False:前のコントロール</param>
		''' <param name="isCursor">カーソルキーによるフォーカス移動か否か</param>
		''' <returns>コントロール</returns>
		Public Function GetNextControl(
			ByRef control As System.Windows.Forms.Control,
			Optional ByVal forward As Boolean = True,
			Optional ByVal isCursor As Boolean = False) _
			As System.Windows.Forms.Control

			Try
				Dim name As String = control.Name
				Dim nextControl As System.Windows.Forms.Control =
					GetNextControl(name, forward, isCursor).Control

				If nextControl.Visible AndAlso nextControl.Enabled Then
					Return nextControl
				End If

				' 非表示 or 非活性の場合フォーカスしないのでフォーカスできるコントロールを探す
				Dim nextName As String = nextControl.Name

				For Each c As TabOrderModel In _modelList
					nextControl = GetNextControl(nextName, forward, isCursor).Control
					If nextControl.Visible AndAlso
					   nextControl.Enabled Then
						Return nextControl
					End If
					nextName = nextControl.Name
				Next

				' 全て非表示・非活性なのでアクティブコントロールを返す
				Return control
			Catch ex As System.Exception
				' For Fail-Safe
				System.Diagnostics.Debug.WriteLine(ex.ToString())
				Return Nothing
			End Try

		End Function

		''' <summary>
		''' 内部情報を更新する
		''' </summary>
		''' <param name="form">フォーム</param>
		Public Sub Update(
			ByRef form As System.Windows.Forms.Control)

			_modelList = New System.Collections.Generic.List(Of TabOrderModel)()

			Try
				CreateModelList(form)
			Catch ex As System.Exception
				' For Fail-Safe
				System.Diagnostics.Debug.WriteLine(ex.ToString())
				_modelList = Nothing
				Return
			End Try

#If DEBUG Then
			For Each c As TabOrderModel In _modelList
				System.Diagnostics.Debug.WriteLine(c.ToString())
			Next
#End If
		End Sub

		''' <summary>
		''' モデルリストを作成する
		''' </summary>
		''' <param name="rootControl">ルートコントロール</param>
		Private Sub CreateModelList(
			ByRef rootControl As System.Windows.Forms.Control)

			' ルートコントロール配下の全コントロールを調べる
			For Each item As System.Windows.Forms.Control In GetAllControls(rootControl)
				' フォーカスが当たらないコントロールは無視
				If Not IsTargetControl(item) Then
					Continue For
				End If

				Dim model As New TabOrderModel(item)
				_modelList.Add(model)
			Next

			' 内部的にナンバリングした重複無しのタブインデックス値を設定する
			UpdateUniqueTabIndex()

			' 前後のコントロールを設定する
			UpdatePrevNextControl()

		End Sub

		''' <summary>
		''' ルートコントロール配下の全コントロールの一覧を返す
		''' </summary>
		''' <param name="rootControl">ルートコントロール</param>
		''' <returns>全コントロールの一覧</returns>
		Private Iterator Function GetAllControls(
			rootControl As System.Windows.Forms.Control) _
			As System.Collections.Generic.IEnumerable(Of System.Windows.Forms.Control)

			For Each c As System.Windows.Forms.Control In rootControl.Controls
				Yield c
				For Each a As System.Windows.Forms.Control In GetAllControls(c)
					Yield a
				Next
			Next

		End Function

		''' <summary>
		''' タブオーダーの対象コントロールかどうかを返す
		''' </summary>
		''' <param name="target">対象コントロール</param>
		''' <returns>True:タブオーダーの対象, False:タブオーダーの対象外</returns>
		Private Function IsTargetControl(
			ByRef target As System.Windows.Forms.Control) _
			As Boolean

			If target.TabStop = False OrElse
			   TypeOf target Is System.Windows.Forms.Panel OrElse
			   TypeOf target Is System.Windows.Forms.GroupBox Then
				Return False
			End If
			Return True

		End Function

		''' <summary>
		''' 内部的にナンバリングした重複無しのタブインデックス値を設定する
		''' </summary>
		Private Sub UpdateUniqueTabIndex()

			_modelList.Sort(New SortHelperOfHierarchicalTabIndices(Sort.Asc))

			Dim index As Integer = 0

			_modelList = _modelList.Where(Function(x) Not x.IsUserControlChild) _
								   .OrderBy(Function(x) x.LastIndex) _
								   .ToList()

			For i As Integer = 0 To _modelList.Count - 1
				Dim model = _modelList(i)

				If Not model.IsUserControlChild Then
					' ユーザーコントロールの子供以外は無条件に設定
					model.UniqueTabIndex = index
					index += 1
					Continue For
				End If
			Next i

		End Sub

		''' <summary>
		''' 前後のコントロールを設定する
		''' </summary>
		Private Sub UpdatePrevNextControl()

			For Each model As TabOrderModel In _modelList

				' シンプルに次(or前)のユニークタブインデックスのコントロールを設定する
				For i As Integer = 0 To 2 - 1 ' 2はforward=True/Falseを表す

					Dim forward As Boolean = If((i = 0), True, False)
					Dim targetIndex As Integer = If(forward, model.UniqueTabIndex + 1, model.UniqueTabIndex - 1)

					Dim updateModel As TabOrderModel = Nothing
					Dim foundModel As TabOrderModel = Nothing

					' モデルリストからターゲットと一致するインデックスを探す
					foundModel = _modelList.FirstOrDefault(Function(x) x.UniqueTabIndex = targetIndex)

					If foundModel Is Nothing Then
						' ターゲットが見つからないのでリストの先頭or末尾からインデックスを取得する

						If forward Then
							' Nextの場合は昇順ソートして先頭から検索
							foundModel = _modelList.OrderBy(Function(x) x.UniqueTabIndex) _
												   .FirstOrDefault(Function(x) x.UniqueTabIndex >= 0)
						Else
							' Prevの場合は降順ソートして先頭から検索
							foundModel = _modelList.OrderByDescending(Function(x) x.UniqueTabIndex) _
												   .FirstOrDefault(Function(x) x.UniqueTabIndex >= 0)
						End If

						If foundModel Is Nothing Then
							' 有効な値が見つからない
							Throw New ControlNotFoundException($"Next or Preview Control not found. Info=[{model}]")
						End If
					End If

					updateModel = New TabOrderModel(foundModel.Control)
					updateModel.UniqueTabIndex = foundModel.UniqueTabIndex

					If forward Then
						model.NextControl = updateModel
					Else
						model.PrevControl = updateModel
					End If

				Next i
			Next

		End Sub

		''' <summary>
		''' 次or前のコントロールをモデルリストから検索する
		''' </summary>
		''' <param name="name">検索対象のコントロール名</param>
		''' <param name="forward">検索方向</param>
		''' <param name="isCursor">カーソルキーによるフォーカス移動か否か</param>
		''' <returns>次or前のコントロール</returns>
		Private Function GetNextControl(
			ByVal name As String,
			ByVal forward As Boolean,
			ByVal isCursor As Boolean) _
			As TabOrderModel

			Dim target As TabOrderModel =
				_modelList.FirstOrDefault(Function(x) x.Control.Name = name)

			If target Is Nothing Then
				' 有効な値が見つからない
				Throw New ControlNotFoundException($"Control not found. Info=[{name}]")
			End If

			Dim ret As TabOrderModel = If(forward, target.NextControl, target.PrevControl)

			If Not isCursor Then
				' TABキーによるフォーカス移動

				If Not ret.IsRadioButton Then
					' 次or前のコントロールがラジオボタン以外であればそのまま返す
					Return ret
				Else
					' 次or前のコントロールがラジオボタンの場合

					' 同じグループのラジオボタンが全て未チェックか調べて真であればチェックをつけて返す
					If ret.Control.Visible AndAlso ret.Control.Enabled AndAlso
					   IsAllUncheckedByRadioButton(ret) Then
						Dim rb As System.Windows.Forms.RadioButton =
						   DirectCast(ret.Control, System.Windows.Forms.RadioButton)
						rb.Checked = True
						Return ret
					End If

					Dim nextControl As TabOrderModel = ret

					' 同じグループのラジオボタンのいずれかにチェックがついている場合
					For Each c As TabOrderModel In _modelList

						' ラジオボタン以外 or チェック済の場合はそのまま返す
						If Not nextControl.IsRadioButton OrElse
						   IsCheckedByRadioButton(nextControl) Then
							Return nextControl
						End If

						' 未チェックのラジオボタンの場合、次or前のコントロールを取得する
						nextControl = GetNextControl(nextControl, forward)
					Next

					Return nextControl

				End If

			Else
				' カーソルキーによるフォーカス移動(同階層のみに制限する)

				If target.ParentLastIndex = ret.ParentLastIndex Then
					' 次or前のコントロールが同階層であればそのまま返す
					Return ret
				Else
					' 次or前のコントロールが同階層でない場合は同階層のコントロールを探して返す

					Dim nextControl As TabOrderModel = ret

					For Each c As TabOrderModel In _modelList

						If target.ParentLastIndex = nextControl.ParentLastIndex Then
							Return nextControl
						End If

						nextControl = GetNextControlBySameLevel(target, forward)
					Next

					Return nextControl

				End If
			End If

		End Function

		''' <summary>
		''' 同じグループのラジオボタンが全て未チェックか否かを返す
		''' </summary>
		''' <param name="target">対象のラジオボタン</param>
		''' <returns>True:全て未チェック, False:チェック済</returns>
		Private Function IsAllUncheckedByRadioButton(
			target As TabOrderModel) _
			As Boolean

			Dim rbIsAllUnchecked As Boolean =
				_modelList.Where(Function(x) x.IsRadioButton AndAlso
											 target.IsRadioButton AndAlso
											 (x.ParentLastIndex = target.ParentLastIndex)) _
						  .All(Function(x)
								   Dim rb As System.Windows.Forms.RadioButton =
										DirectCast(x.Control, System.Windows.Forms.RadioButton)
								   Return Not rb.Checked
							   End Function)

			Return rbIsAllUnchecked

		End Function

		''' <summary>
		''' ラジオボタンがチェックされているかを返す
		''' </summary>
		''' <param name="target">対象のラジオボタン</param>
		''' <returns>True:チェック, False:未チェック</returns>
		Private Function IsCheckedByRadioButton(
			target As TabOrderModel) _
			As Boolean

			Dim rb As System.Windows.Forms.RadioButton =
				DirectCast(target.Control, System.Windows.Forms.RadioButton)
			Return rb.Checked

		End Function

		''' <summary>
		''' 次or前のコントロールを取得する
		''' </summary>
		''' <param name="target">対象コントロール</param>
		''' <param name="forward">検索方向</param>
		''' <returns>次or前のコントロール</returns>
		Private Function GetNextControl(
			target As TabOrderModel,
			ByVal forward As Boolean) _
			As TabOrderModel

			Dim ret As TabOrderModel

			If forward Then
				ret = _modelList.FirstOrDefault(Function(x) x.UniqueTabIndex > target.UniqueTabIndex)
				If ret Is Nothing Then
					ret = _modelList.First()
				End If
			Else
				ret = _modelList.LastOrDefault(Function(x) x.UniqueTabIndex < target.UniqueTabIndex)
				If ret Is Nothing Then
					ret = _modelList.Last()
				End If
			End If

			Return ret

		End Function

		''' <summary>
		''' 対象コントロールと同階層の次or前のコントロールを取得する
		''' </summary>
		''' <param name="target">対象コントロール</param>
		''' <param name="forward">検索方向</param>
		''' <returns>同階層の次or前のコントロール</returns>
		Private Function GetNextControlBySameLevel(
			target As TabOrderModel,
			ByVal forward As Boolean) _
			As TabOrderModel

			Dim ret As TabOrderModel

			If forward Then
				ret = _modelList.FirstOrDefault(Function(x) x.UniqueTabIndex > target.UniqueTabIndex AndAlso
															x.ParentLastIndex = target.ParentLastIndex)
				If ret Is Nothing Then
					ret = _modelList.FirstOrDefault(Function(x) x.ParentLastIndex = target.ParentLastIndex)
				End If
			Else
				ret = _modelList.LastOrDefault(Function(x) x.UniqueTabIndex < target.UniqueTabIndex AndAlso
														   x.ParentLastIndex = target.ParentLastIndex)
				If ret Is Nothing Then
					ret = _modelList.LastOrDefault(Function(x) x.ParentLastIndex = target.ParentLastIndex)
				End If
			End If

			If ret Is Nothing Then
				ret = target
			End If

			Return ret

		End Function

	End Class
End Namespace