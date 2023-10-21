Imports System.Runtime.InteropServices

''' <summary>
''' コピー/カット操作をフックして、クリップボードの内容を改変するカスタムテキストボックス
''' </summary>
Public Class TextBoxEx
    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Private Shared Function SetWindowLong(hWnd As IntPtr, nIndex As Integer, newWndProc As IntPtr) As IntPtr
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Private Shared Function CallWindowProc(wndProc As IntPtr, hWnd As IntPtr, msg As UInteger, wParam As IntPtr, lParam As IntPtr) As IntPtr
    End Function

    Private Delegate Function WndProcCallbackDelegate(hWnd As IntPtr, msg As UInteger, wParam As IntPtr, lParam As IntPtr) As IntPtr

    Private Const GWL_WNDPROC As Integer = -4   'ウィンドウプロシージャの新しいアドレスを設定します
    Private Const WM_CUT As Integer = &H300     'ウィンドウメッセージID(カット)
    Private Const WM_COPY As Integer = &H301    'ウィンドウメッセージID(コピー)

    Private _myWndHandle As IntPtr
    Private _isStartedWndProc As Boolean = False

    Public Sub New()
        InitializeComponent()

        'If DesignMode Then
        '    Return     '効果がなかった...orz
        'End If
    End Sub

    ''' <summary>
    ''' Enterイベントハンドラ
    ''' </summary>
    ''' <param name="e">イベント引数</param>
    Protected Overrides Sub OnEnter(e As EventArgs)
        If Not _isStartedWndProc Then
            'コンストラクタでStartWndProc()を呼ぶとVisualStudioのデザイナが落ちるのでここで呼ぶ
            _isStartedWndProc = True
            StartWndProc()
        End If

        MyBase.OnEnter(e)
    End Sub

    Private Sub StartWndProc()
        'コントロールのウィンドウプロシージャをフックする
        _myWndHandle = SetWindowLong(
                            MyBase.Handle,
                            GWL_WNDPROC,
                            Marshal.GetFunctionPointerForDelegate(New WndProcCallbackDelegate(AddressOf WndProcCallback)))
    End Sub

    ''' <summary>
    ''' ウィンドウに送信されたメッセージを処理する、アプリケーションで定義するコールバック関数
    ''' </summary>
    ''' <param name="hWnd">ウィンドウへのハンドル</param>
    ''' <param name="msg">メッセージ</param>
    ''' <param name="wParam">追加のメッセージ情報</param>
    ''' <param name="lParam">追加のメッセージ情報</param>
    ''' <returns>メッセージ処理の結果</returns>
    Private Function WndProcCallback(hWnd As IntPtr, msg As UInteger, wParam As IntPtr, lParam As IntPtr) As IntPtr
        If msg = WM_COPY Then
            Clipboard.SetText(MyBase.Text & "_COPY")    '改変
            Return IntPtr.Zero ' デフォルトのコピー操作をキャンセル
        End If

        If msg = WM_CUT Then
            Clipboard.SetText(MyBase.SelectedText & "_CUT")    '改変
            MyBase.SelectedText = String.Empty
            Return IntPtr.Zero ' デフォルトのカット操作をキャンセル
        End If

        ' 上記以外のメッセージはデフォルトのウィンドウプロシージャに移譲
        Return CallWindowProc(_myWndHandle, hWnd, msg, wParam, lParam)
    End Function
End Class
