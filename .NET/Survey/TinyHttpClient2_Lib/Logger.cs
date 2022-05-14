using System;
using System.IO;
using System.Threading;

namespace CommonLib
{
    public static class Constant
    {
        /// <summary>日時フォーマット（yyyy/MM/dd hh:mm:ss.fff）</summary>
        public static readonly string FORMAT_YYYYMMDDHHMMSSFFF = "yyyy/MM/dd HH:mm:ss.fff";

        public static readonly string FORMAT_EXCEPTION = "\nMessage={0}, StackTrace={1}";

        public static readonly char TAB = '\t';

        public static readonly char COMMA = ',';
    }

    /// <summary>
    /// ロガークラス
    /// </summary>
    public static class Logger
    {
        /// <summary>
        /// 出力フォーマット
        /// {0}:出力時間
        /// {1}:スレッドID
        /// {2}:ログの内容
        /// </summary>
        private static readonly string FORMAT = "{0}[{1}]:{2}";

        /// <summary>
        /// 出力フォーマット（例外用）
        /// {0}:例外メッセージ
        /// {1}:スタックトレース
        /// </summary>
        private static readonly string FORMAT_EX = "Message={0}, Stack={1}";

        /// <summary>ライター</summary>
        private static StreamWriter _sw = null;

        /// <summary>
        /// 初期化する
        /// </summary>
        /// <param name="path">ログファイルパス</param>
        public static void Initialize(string path)
        {
            Dispose();

            _sw = new StreamWriter(path, true);
        }

        /// <summary>
        /// ログを出力する
        /// </summary>
        /// <param name="data">出力内容</param>
        public static void Write(string data)
        {
            if (_sw == null) return;

            var writeData = string.Format(
                                    FORMAT,
                                    DateTime.Now.ToString(Constant.FORMAT_YYYYMMDDHHMMSSFFF),
                                    Thread.CurrentThread.ManagedThreadId,
                                    data);
            try
            {
                _sw.WriteLine(writeData);
                _sw.Flush();
            }
            catch
            {
                // 無視
            }
        }

        /// <summary>
        /// ログを出力する（例外用）
        /// </summary>
        /// <param name="ex">例外オブジェクト</param>
        public static void Write(Exception ex)
        {
            if (ex == null) return;
            var message = string.Format(FORMAT_EX, ex.Message, ex.StackTrace);
            Write(message);
        }

        /// <summary>
        /// ログ出力を終了する
        /// </summary>
        public static void Dispose()
        {
            if (_sw != null)
            {
                _sw.Dispose();
                _sw = null;
            }
        }
    }
}
