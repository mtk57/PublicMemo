/*

http/httpsの接続テストアプリ (.NET4.5～)

.NET3.5をベースにしたかったので、あえてHttpWebRequestを使用している。

使用例.
1.XAMPPでApacheを開始する。 
　⇒XAMPPは TLSv1～v1.2まで対応している。ポートは4433。
2.WiresharkでNpcap LoopBack Adapterのパケットをキャプチャを開始する。
　⇒Npcap以外だとループバックがキャプチャできない
3.本Appを起動し、TLSバージョンを指定して、GET or POSTを実行する。
　⇒TLSバージョン未指定時はTLSv1となる
4.Wiresharkで「TLS」でフィルタして確認する。

*/

using System;
using System.Windows.Forms;

namespace TinyHttpClient
{
    static class Program
    {
        /// <summary>
        /// アプリケーションのメイン エントリ ポイントです。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new FormMain());
        }
    }
}
