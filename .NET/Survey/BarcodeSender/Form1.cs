using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace BarcodeSender
{
    public partial class Form1 : Form
    {
        [DllImport( "user32.dll" )]
        public static extern IntPtr FindWindow ( string lpClassName, string lpWindowName );

        [DllImport( "user32.dll" )]
        public static extern bool SetForegroundWindow ( IntPtr hWnd );

        [DllImport( "user32.dll" )]
        public static extern int SendMessage ( IntPtr hWnd, int Msg, int wParam, int lParam );

        [DllImport( "user32.dll" )]
        public static extern bool PostMessage ( IntPtr hWnd, uint Msg, int wParam, int lParam );

        private const int WM_KEYDOWN = 0x0100;
        private const int WM_KEYUP = 0x0101;
        private const int WM_CHAR = 0x0102;

        public Form1 ()
        {
            InitializeComponent();
        }

        private void Form1_Load ( object sender, EventArgs e )
        {
            textBoxReceiveWindowCaption.Text = "Barcode Receiver";
        }

        private void buttonSend_Click ( object sender, EventArgs e )
        {
            EmulateBarcodeScan( textBoxSendData.Text, textBoxReceiveWindowCaption.Text );
        }


        private void EmulateBarcodeScan ( string barcode, string targetWindowTitle )
        {
            IntPtr hWnd = FindWindow( null, targetWindowTitle );
            if ( hWnd != IntPtr.Zero )
            {
                SetForegroundWindow( hWnd );
                System.Threading.Thread.Sleep( 100 ); // ウィンドウがフォアグラウンドになるのを待つ

                foreach ( char c in barcode )
                {
                    PostMessage( hWnd, WM_CHAR, c, 0 );
                    System.Threading.Thread.Sleep( 10 ); // 各文字入力に少し遅延を加える
                }

                // エンターキーを送信
                PostMessage( hWnd, WM_KEYDOWN, 0x0D, 0 );
                PostMessage( hWnd, WM_KEYUP, 0x0D, 0 );
            }
            else
            {
                MessageBox.Show( $"ターゲットウィンドウ '{targetWindowTitle}' が見つかりません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error );
            }
        }
    }
}
