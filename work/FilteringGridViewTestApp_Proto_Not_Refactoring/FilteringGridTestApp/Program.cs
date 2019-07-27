using System;
using System.Threading;
using System.Windows.Forms;

namespace FilteringGridTestApp
{
    static class Program
    {
        /// <summary>
        /// アプリケーションのメイン エントリ ポイントです。
        /// </summary>
        [STAThread]
        static void Main ()
        {
            Application.ThreadException += new ThreadExceptionEventHandler( Application_ThreadException );
            AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler( CurrentDomain_UnhandledException );

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault( false );
            Application.Run( new Form1() );
        }

        private static void Application_ThreadException ( object sender, ThreadExceptionEventArgs e ) => Utils.ShowError( Utils.GetExceptionMessage( e.Exception ) );

        private static void CurrentDomain_UnhandledException ( object sender, UnhandledExceptionEventArgs e ) => Utils.ShowError( Utils.GetExceptionMessage( ( Exception ) e.ExceptionObject ) );
    }
}
