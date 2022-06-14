using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;

namespace TextCompressor.Common
{
    /// <summary>
    /// ユーティリティー
    /// </summary>
    internal class Utils
    {
        public static List<string> SplitNewLine(string data)
        {
            return new List<string>(data.Replace("\r\n", "\n").Split(new[] { '\n', '\r' }));
        }

        //------------------------------------------
        //文字列を圧縮BASE64に変換して返す
        //https://gurizuri0505.halfmoon.jp/20121211/52033
        //------------------------------------------
        public static string Base64FromStringComp(string inStr)
        {
            // 文字列をバイト配列に変換します 
            var src = Encoding.UTF8.GetBytes(inStr);

            // 入出力用のストリームを生成します 
            var ms = new MemoryStream();
            var ds = new DeflateStream(ms, CompressionMode.Compress, true);

            // ストリームに圧縮するデータを書き込みます 
            ds.Write(src, 0, src.Length);
            ds.Close();

            // 圧縮されたデータを バイト配列で取得します 
            var dst = ms.ToArray();

            //Base64で文字列に変換
            return Convert.ToBase64String(dst, Base64FormattingOptions.InsertLineBreaks);
        }

        //------------------------------------------
        //BASE64文字列を戻し解凍の上で文字列に変換して返す
        //https://gurizuri0505.halfmoon.jp/20121211/52033
        //------------------------------------------
        public static string StringFromBase64Comp(string base64Str)
        {
            var bs = Convert.FromBase64String(base64Str);

            // 入出力用のストリームを生成します 
            var ms = new MemoryStream(bs);
            var ms2 = new MemoryStream();
            var ds = new DeflateStream(ms, CompressionMode.Decompress);

            // MemoryStream に展開します 
            while (true)
            {
                var rb = ds.ReadByte();
                // 読み終わったとき while 処理を抜けます 
                if (rb == -1)
                {
                    break;
                }
                // メモリに展開したデータを読み込みます 
                ms2.WriteByte((byte)rb);
            }

            return Encoding.UTF8.GetString(ms2.ToArray());
        }

        /// <summary>
        /// 文字列を暗号化する
        /// https://dobon.net/vb/dotnet/string/encryptstring.html
        /// </summary>
        /// <param name="srcStr">暗号化する文字列</param>
        /// <param name="password">暗号化に使用するパスワード</param>
        /// <returns>暗号化された文字列</returns>
        public static string EncryptString(string srcStr, string password)
        {
            //RijndaelManagedオブジェクトを作成
            var rijndael = new RijndaelManaged();

            //パスワードから共有キーと初期化ベクタを作成
            byte[] key, iv;

            GenerateKeyFromPassword(password, rijndael.KeySize, out key, rijndael.BlockSize, out iv);

            rijndael.Key = key;
            rijndael.IV = iv;

            //文字列をバイト型配列に変換する
            var strBytes = Encoding.UTF8.GetBytes(srcStr);

            //対称暗号化オブジェクトの作成
            using (var encryptor = rijndael.CreateEncryptor())
            {
                //バイト型配列を暗号化する
                var encBytes = encryptor.TransformFinalBlock(strBytes, 0, strBytes.Length);

                //バイト型配列を文字列に変換して返す
                return Convert.ToBase64String(encBytes);
            }
        }

        /// <summary>
        /// 暗号化された文字列を復号化する
        /// https://dobon.net/vb/dotnet/string/encryptstring.html
        /// </summary>
        /// <param name="srcStr">暗号化された文字列</param>
        /// <param name="password">暗号化に使用したパスワード</param>
        /// <returns>復号化された文字列</returns>
        public static string DecryptString(string srcStr, string password)
        {
            //RijndaelManagedオブジェクトを作成
            var rijndael = new RijndaelManaged();

            //パスワードから共有キーと初期化ベクタを作成
            byte[] key, iv;

            GenerateKeyFromPassword(password, rijndael.KeySize, out key, rijndael.BlockSize, out iv);
            rijndael.Key = key;
            rijndael.IV = iv;

            //文字列をバイト型配列に戻す
            var strBytes = Convert.FromBase64String(srcStr);

            //対称暗号化オブジェクトの作成
            using (var decryptor = rijndael.CreateDecryptor())
            {
                //バイト型配列を復号化する
                //復号化に失敗すると例外CryptographicExceptionが発生
                var decBytes = decryptor.TransformFinalBlock(strBytes, 0, strBytes.Length);

                //バイト型配列を文字列に戻して返す
                return Encoding.UTF8.GetString(decBytes);
            }
        }

        /// <summary>
        /// パスワードから共有キーと初期化ベクタを生成する
        /// https://dobon.net/vb/dotnet/string/encryptstring.html
        /// </summary>
        /// <param name="password">基になるパスワード</param>
        /// <param name="keySize">共有キーのサイズ（ビット）</param>
        /// <param name="key">作成された共有キー</param>
        /// <param name="blockSize">初期化ベクタのサイズ（ビット）</param>
        /// <param name="iv">作成された初期化ベクタ</param>
        private static void GenerateKeyFromPassword(string password,
            int keySize, out byte[] key, int blockSize, out byte[] iv)
        {
            //パスワードから共有キーと初期化ベクタを作成する
            //saltを決める
            byte[] salt = System.Text.Encoding.UTF8.GetBytes("saltは必ず8バイト以上");
            //Rfc2898DeriveBytesオブジェクトを作成する
            System.Security.Cryptography.Rfc2898DeriveBytes deriveBytes =
                new System.Security.Cryptography.Rfc2898DeriveBytes(password, salt);
            //.NET Framework 1.1以下の時は、PasswordDeriveBytesを使用する
            //System.Security.Cryptography.PasswordDeriveBytes deriveBytes =
            //    new System.Security.Cryptography.PasswordDeriveBytes(password, salt);
            //反復処理回数を指定する デフォルトで1000回
            deriveBytes.IterationCount = 1000;

            //共有キーと初期化ベクタを生成する
            key = deriveBytes.GetBytes(keySize / 8);
            iv = deriveBytes.GetBytes(blockSize / 8);
        }

        public static string ReadFile(string filePath)
        {
            using (var sr = new StreamReader(filePath))
            {
                return sr.ReadToEnd();
            }
        }

        public static void WriteFile(string filePath, string writeData)
        {
            using (var sw = new StreamWriter(filePath, false, Encoding.UTF8))
            {
                sw.Write(writeData);
            }
        }

        public static IEnumerable<string> GetFileList(string rootDir, List<string> extensions)
        {
            return FilterFileListByExtensions(GetAllFileList(rootDir), extensions);
        }

        private static IEnumerable<string> GetAllFileList(string rootDir)
        {
            return Directory.EnumerateFiles(rootDir, "*", SearchOption.AllDirectories);
        }

        private static IEnumerable<string> FilterFileListByExtensions(IEnumerable<string> fileList, List<string> extensions)
        {
            var ret = new List<string>();

            foreach(var file in fileList)
            {
                foreach (var ext in extensions)
                {
                    var fileExt = Path.GetExtension(file);

                    if (fileExt == "." + ext)
                    {
                        ret.Add(file);
                    }
                }
            }

            return ret;
        }


        /// <summary>
        /// 辞書を文字列に変換する。
        /// Ex.
        ///   IN:
        ///     key="KEY1", value="VALUE1"
        ///     key="KEY2", value="VALUE2"
        ///   OUT:
        ///     "[KEY1]=[VALUE1],[KEY2]=[VALUE2],"
        /// </summary>
        /// <param name="dict">辞書</param>
        /// <returns>文字列</returns>
        public static string GetDicString(Dictionary<string, string> dict)
        {
            if (dict == null)
            {
                return string.Empty;
            }

            var sb = new StringBuilder();

            foreach (var key in dict.Keys)
            {
                sb.Append(string.Format("[{0}]=[{1}],", key, dict[key]));
            }

            return sb.ToString();
        }

        /// <summary>
        /// .NETのバージョンを返す。
        /// </summary>
        /// <returns>.NETのバージョン</returns>
        public static string GetDotNetVersion()
        {
            return typeof(string).Assembly.GetName().Version.ToString();
        }

        /// <summary>
        /// 自身のディレクトリー名を返す。
        /// </summary>
        /// <returns></returns>
        public static string GetMyDir()
        {
            return Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        }

        /// <summary>
        /// 自身のアセンブリバージョンを返す。
        /// </summary>
        /// <returns>自身のアセンブリバージョン</returns>
        public static string GetMyVersion()
        {
            var asm = Assembly.GetExecutingAssembly().GetName();

            return asm.Version.ToString();
        }
    }
}
