using System;
using System.IO;
using System.Reflection;
using System.Threading;

namespace Grep.Common
{
    internal class Logger
    {
        private const string LOG_NAME = "Grep.log";

        private const string FORMAT_YYYYMMDDHHMMSSFFF = "yyyy/MM/dd HH:mm:ss.fff";

        private const string FORMAT = "{0}[{1}]:{2}";

        private const string FORMAT_EX = "Message={0}, Stack={1}";

        private static StreamWriter _sw = null;

        public static bool IsInitSuccess
        {
            get { return _sw != null; }
        }

        public static void Initialize(string path = "")
        {
            if (string.IsNullOrEmpty(path))
            {
                path = Path.Combine(GetMyDir(), LOG_NAME);
            }

            Dispose();

            _sw = new StreamWriter(path, true);
        }

        public static void Debug(string writeData)
        {
            Write(writeData);
        }

        public static void Info(string writeData)
        {
            Write(writeData);
        }

        public static void Warn(string writeData)
        {
            Write(writeData);
        }

        public static void Error(string writeData)
        {
            Write(writeData);
        }

        private static void Write(string data)
        {
            if (_sw == null) return;

            var writeData = string.Format(
                                    FORMAT,
                                    DateTime.Now.ToString(FORMAT_YYYYMMDDHHMMSSFFF),
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

        public static void Dispose()
        {
            if (_sw != null)
            {
                _sw.Dispose();
                _sw = null;
            }
        }

        private static string GetMyDir()
        {
            return Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        }
    }
}
