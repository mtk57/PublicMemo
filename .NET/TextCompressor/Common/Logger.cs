using System;
using System.IO;
using System.Threading;

namespace TextCompressor.Common
{
    /// <summary>
    /// ロガー
    /// </summary>
    /// <remarks>
    /// 簡易的なロガーなので、将来的にはLog4Net等に置き換えることを想定している。
    /// ログ書き込みメソッドが4つあるがどれも同じことをしている。
    /// </remarks>
    internal class Logger
    {
        /// <summary>
        /// 日付フォーマット
        /// </summary>
        private const string FORMAT_YYYYMMDDHHMMSSFFF = "yyyy/MM/dd HH:mm:ss.fff";

        /// <summary>
        /// ログフォーマット
        /// </summary>
        private const string FORMAT = "{0}[{1}]:{2}";

        /// <summary>
        /// 例外情報フォーマット
        /// </summary>
        private const string FORMAT_EX = "Message={0}, Stack={1}";

        /// <summary>
        /// StreamWriter
        /// </summary>
        private static StreamWriter sw = null;

        /// <summary>
        /// 初期化成功有無
        /// </summary>
        public static bool IsInitSuccess
        {
            get
            {
                return sw != null;
            }
        }

        /// <summary>
        /// 初期化する。
        /// </summary>
        /// <param name="filePath">ログファイルパス</param>
        public static void Initialize(string filePath = "")
        {
            if (string.IsNullOrEmpty(filePath))
            {
                filePath = Path.Combine(Utils.GetMyDir(), "TextCompressor.log");
            }

            Dispose();

            sw = new StreamWriter(filePath, true);
        }

        /// <summary>
        /// ログを書き込む。
        /// </summary>
        /// <param name="writeData">書き込むデータ</param>
        public static void Debug(string writeData)
        {
            Write(writeData);
        }

        /// <summary>
        /// ログを書き込む。
        /// </summary>
        /// <param name="writeData">書き込むデータ</param>
        public static void Info(string writeData)
        {
            Write(writeData);
        }

        /// <summary>
        /// ログを書き込む。
        /// </summary>
        /// <param name="writeData">書き込むデータ</param>
        public static void Warn(string writeData)
        {
            Write(writeData);
        }

        /// <summary>
        /// ログを書き込む。
        /// </summary>
        /// <param name="writeData">書き込むデータ</param>
        public static void Error(string writeData)
        {
            Write(writeData);
        }

        /// <summary>
        /// ログを書き込む。
        /// </summary>
        /// <param name="writeData">書き込むデータ</param>
        private static void Write(string writeData)
        {
            if (sw == null)
            {
                return;
            }

            var data = string.Format(FORMAT, DateTime.Now.ToString(FORMAT_YYYYMMDDHHMMSSFFF),
                            Thread.CurrentThread.ManagedThreadId, writeData);
            try
            {
                sw.WriteLine(data);

                sw.Flush();
            }
            catch
            {
                // 無視
            }
        }

        /// <summary>
        /// 後処理をする。
        /// </summary>
        public static void Dispose()
        {
            if (sw != null)
            {
                sw.Dispose();

                sw = null;
            }
        }
    }
}
