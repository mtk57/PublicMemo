Imports System.Runtime.InteropServices

Namespace Utility
    'http://wisdom.sakura.ne.jp/system/winapi/win32/win92.html
    'https://lets-csharp.com/how-to-clipboard-listener/
    'https://qiita.com/kob58im/items/2697ea4c12c72ecd86c8
    'https://gist.github.com/unarist/6342758


    'Friend Class ClipboardHelper
    '    Inherits System.Windows.Forms.Form

    '    Public Event UpdateClipboard As EventHandler

    '    <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True)>
    '    Private Shared Sub AddClipboardFormatListener(hwnd As System.IntPtr)
    '    End Sub

    '    <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True)>
    '    Private Shared Sub RemoveClipboardFormatListener(hwnd As System.IntPtr)
    '    End Sub

    '    Private Const WM_CLIPBOARDUPDATE As Integer = &H31D

    '    Private _control As System.Windows.Forms.Control

    '    'Public Sub New(c As System.Windows.Forms.Control)
    '    '    _control = c
    '    'End Sub

    '    'Protected Overrides Sub OnCreateControl()
    '    '    AddClipboardFormatListener(_control.Handle)
    '    '    MyBase.OnCreateControl()
    '    'End Sub

    '    'Protected Overrides Sub Dispose(disposing As Boolean)
    '    '    RemoveClipboardFormatListener(_control.Handle)
    '    '    MyBase.Dispose(disposing)
    '    'End Sub

    '    'Public Sub New(ByVal c As System.Windows.Forms.Control)
    '    '    AddHandler c.HandleCreated, AddressOf OnHandleCreated
    '    '    AddHandler c.HandleDestroyed, AddressOf OnHandleDestroyed
    '    'End Sub

    '    'Private Sub OnHandleCreated(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    '    AddClipboardFormatListener(Handle)
    '    'End Sub

    '    'Private Sub OnHandleDestroyed(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    '    RemoveClipboardFormatListener(Handle)
    '    'End Sub

    '    Protected Overrides Sub OnLoad(e As EventArgs)
    '        AddClipboardFormatListener(Handle)
    '        MyBase.OnLoad(e)
    '    End Sub

    '    Protected Overrides Sub Dispose(disposing As Boolean)
    '        RemoveClipboardFormatListener(Handle)
    '        MyBase.Dispose(disposing)
    '    End Sub

    '    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
    '        If m.Msg = WM_CLIPBOARDUPDATE AndAlso System.Windows.Forms.Clipboard.ContainsText() Then
    '            RaiseEvent UpdateClipboard(Me, New System.EventArgs())
    '            m.Result = System.IntPtr.Zero
    '        Else
    '            MyBase.WndProc(m)
    '        End If
    '    End Sub

    'End Class

    <System.Security.Permissions.PermissionSet(System.Security.Permissions.SecurityAction.Demand, Name:="FullTrust")>
    Friend Class ClipboardHelper
        Inherits System.Windows.Forms.NativeWindow

        Public Event UpdateClipboard As System.EventHandler

        'クリップボードビュワーにウィンドウを登録する
        '戻り値は、このビュワーの次に登録されているビュワーのハンドルです
        '
        'この戻り値の値が意味するものは、クリップボードビュワーの作成に重要です
        '実は、Windows は1つのビュワーにしかメッセージを通知しません
        'つまり、最後に SetClipboardViewer() を用いて登録したビュワーです
        '
        'しかし、一度に複数のビュワーが実行されることも十分ありえます
        'そのため、プログラムは SetClipboardViewer() が返したウィンドウに
        'WM_DRAWCLIPBOARD メッセージを送信する必要があるのです

        'こうして、上から下へ、鎖にようにビュワープログラムがメッセージを流します
        'このクリップボードビュワーの構造をクリップボードビュワーチェインと呼びます
        Private Declare Auto Function SetClipboardViewer Lib "user32" (hWndNewViewer As System.IntPtr) As System.IntPtr


        Private Declare Auto Function ChangeClipboardChain Lib "user32" (ByVal hWndRemove As System.IntPtr, ByVal hWndNewNext As System.IntPtr) As Boolean

        Private Declare Auto Function SendMessage Lib "user32" (ByVal hWnd As System.IntPtr, ByVal Msg As Integer, ByVal wParam As System.IntPtr, ByVal lParam As System.IntPtr) As Integer

        'クリップボードに変更があるとクリップボードビュワーに通知される
        Private Const WM_DRAWCLIPBOARD As Integer = &H308

        'チェインに参加している全てのウィンドウは、このメッセージを処理する義務があります
        Private Const WM_CHANGECBCHAIN As Integer = &H30D

        Private _nextHandle As System.IntPtr = IntPtr.Zero

        Public Sub New(ByVal c As System.Windows.Forms.Control)
            AddHandler c.HandleCreated, AddressOf OnHandleCreated
            AddHandler c.HandleDestroyed, AddressOf OnHandleDestroyed
        End Sub

        Private Sub OnHandleCreated(ByVal sender As System.Object, ByVal e As System.EventArgs)
            AssignHandle((CType(sender, System.Windows.Forms.Control)).Handle)

            _nextHandle = SetClipboardViewer(Me.Handle)
        End Sub

        Private Sub OnHandleDestroyed(ByVal sender As System.Object, ByVal e As System.EventArgs)
            ChangeClipboardChain(Me.Handle, _nextHandle)
            ReleaseHandle()
        End Sub

        Protected Overrides Sub WndProc(ByRef msg As System.Windows.Forms.Message)
            Select Case msg.Msg
                Case WM_DRAWCLIPBOARD
                    If _nextHandle <> IntPtr.Zero Then
                        SendMessage(_nextHandle, msg.Msg, msg.WParam, msg.LParam)
                    End If

                    If Clipboard.ContainsText() Then
                        RaiseEvent UpdateClipboard(Me, New System.EventArgs())
                    End If

                Case WM_CHANGECBCHAIN
                    If msg.WParam = _nextHandle Then
                        _nextHandle = msg.LParam
                    ElseIf _nextHandle <> IntPtr.Zero Then
                        SendMessage(_nextHandle, msg.Msg, msg.WParam, msg.LParam)
                    End If
            End Select

            MyBase.WndProc(msg)
        End Sub
    End Class
End Namespace


