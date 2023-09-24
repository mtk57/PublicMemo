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
		Private _modelDict As System.Collections.Generic.Dictionary(Of String, TabOrderModel)

		Private Sub New()
			' do nothing
		End Sub

		''' <summary>
		''' コンストラクタ
		''' </summary>
		''' <param name="form">フォーム</param>
		Public Sub New(form As System.Windows.Forms.Control)
			Update(form)
		End Sub

		''' <summary>
		''' カレントコントロールの次(もしくは前)のコントロールを返す
		''' </summary>
		''' <param name="control">カレントコントロール</param>
		''' <param name="forward">True:次のコントロール、False:前のコントロール</param>
		''' <returns>コントロール</returns>
		Public Function GetNextControl(control As System.Windows.Forms.Control, Optional forward As Boolean = True) As System.Windows.Forms.Control
			Dim name As Object = control.Name
			Dim nextControl As Object = If(forward, _modelDict(name).NextControl.Control, _modelDict(name).PrevControl.Control)

			If nextControl.Visible Then
				Return nextControl
			End If

			' 非表示の場合フォーカスしないので表示されているコントロールを探す
			Dim nextName As String = nextControl.Name

			For Each c As TabOrderModel In _modelList
				If forward Then
					If _modelDict(nextName).NextControl.Control.Visible Then
						Return _modelDict(nextName).NextControl.Control
					End If
					nextName = _modelDict(nextName).NextControl.Control.Name
				Else
					If _modelDict(nextName).PrevControl.Control.Visible Then
						Return _modelDict(nextName).PrevControl.Control
					End If
					nextName = _modelDict(nextName).PrevControl.Control.Name
				End If
			Next

			' 全て非表示なのでアクティブコントロールを返す
			Return control
		End Function

		''' <summary>
		''' 内部情報を更新する
		''' </summary>
		''' <param name="form">フォーム</param>
		Public Sub Update(form As System.Windows.Forms.Control)
			_modelList = New System.Collections.Generic.List(Of TabOrderModel)()
			_modelDict = New System.Collections.Generic.Dictionary(Of String, TabOrderModel)()

			CreateModelList(form)
			CreateModelDict()

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
		Private Sub CreateModelList(rootControl As System.Windows.Forms.Control)
			' ルートコントロール配下の全コントロールを調べる
			For Each item As System.Windows.Forms.Control In GetAllControls(rootControl)
				' コンテナ系はフォーカスが当たらないので無視
				If IsContainer(item) Then
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
		Private Iterator Function GetAllControls(rootControl As System.Windows.Forms.Control) As System.Collections.Generic.IEnumerable(Of System.Windows.Forms.Control)
			For Each c As System.Windows.Forms.Control In rootControl.Controls
				Yield c
				For Each a As System.Windows.Forms.Control In GetAllControls(c)
					Yield a
				Next
			Next
		End Function

		''' <summary>
		''' 対象コントロールがコンテナ系かどうかを返す
		''' </summary>
		''' <param name="target">対象コントロール</param>
		''' <returns>True:コンテナ系, False:コンテナ系以外</returns>
		Private Function IsContainer(target As System.Windows.Forms.Control) As Boolean
			If TypeOf target Is System.Windows.Forms.Panel OrElse
			   TypeOf target Is System.Windows.Forms.GroupBox Then
				Return True
			End If
			Return False
		End Function

		''' <summary>
		''' 内部的にナンバリングした重複無しのタブインデックス値を設定する
		''' </summary>
		Private Sub UpdateUniqueTabIndex()
			_modelList.Sort(New SortHelperOfHierarchicalTabIndices(Sort.Asc))

			Dim index As Integer = 0
			Dim groupIndex As Integer? = Nothing

			For i As Integer = 0 To _modelList.Count - 1
				Dim model = _modelList(i)

				If Not model.IsRadioButton Then
					' ラジオボタン以外は無条件に設定
					model.UniqueTabIndex = index
					index += 1
					Continue For
				End If

				' ラジオボタンの場合は同グループの最初のコントロールをタブオーダーの対象とする
				If groupIndex Is Nothing OrElse groupIndex <> model.ParentLastIndex Then
					model.UniqueTabIndex = index
					index += 1
					groupIndex = model.ParentLastIndex
				End If
			Next i
		End Sub


		''' <summary>
		''' 前後のコントロールを設定する
		''' </summary>
		Private Sub UpdatePrevNextControl()
			For Each model As TabOrderModel In _modelList
				If model.UniqueTabIndex >= 0 Then
					' ユニークタブインデックスが設定済の場合は、シンプルに次(or前)のユニークタブインデックスのコントロールを設定する
					UpdatePrevNextControlByModel(model)
				End If
			Next

			For Each model As TabOrderModel In _modelList
				If model.UniqueTabIndex Is Nothing AndAlso model.IsRadioButton Then
					' ユニークタブインデックスが未設定のラジオボタンの場合は、同じユニークタブインデックスのコントロールを設定する
					UpdatePrevNextControlByModelForRadioButton(model)
				End If
			Next
		End Sub

		''' <summary>
		''' 次(or前)のユニークタブインデックスのコントロールを設定する
		''' </summary>
		''' <param name="model">モデル</param>
		''' <exception cref="ControlNotFoundException"></exception>
		Private Sub UpdatePrevNextControlByModel(model As TabOrderModel)
			For i As Integer = 0 To 2 - 1 ' 2はforward=True/Falseを表す
				Dim forward As Boolean = If((i = 0), True, False)
				Dim targetIndex As Integer? = If(forward, model.UniqueTabIndex + 1, model.UniqueTabIndex - 1)

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
						Throw New ControlNotFoundException("Next or Preview Control not found. Info=[{model}]")
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
		End Sub

		''' <summary>
		''' ラジオボタンの場合は、同じユニークタブインデックスのコントロールを設定する
		''' </summary>
		''' <param name="model">モデル</param>
		''' <exception cref="ControlNotFoundException"></exception>
		Private Sub UpdatePrevNextControlByModelForRadioButton(model As TabOrderModel)
			' モデルリストから条件に合致するモデルを探す
			Dim enableRadioButton As TabOrderModel = _modelList.FirstOrDefault(Function(x) x.UniqueTabIndex >= 0 AndAlso
																						   x.ParentLastIndex = model.ParentLastIndex AndAlso
																						   x.IsRadioButton)
			If enableRadioButton Is Nothing Then
				Throw New ControlNotFoundException("Next or Preview Control not found. Info=[{model}]")
			End If

			model.NextControl = enableRadioButton.NextControl
			model.PrevControl = enableRadioButton.PrevControl
		End Sub

		''' <summary>
		''' モデル辞書を作成する
		''' </summary>
		Private Sub CreateModelDict()
			_modelDict = _modelList.ToDictionary(Function(x) x.Control.Name, Function(x) x)
		End Sub
	End Class
End Namespace