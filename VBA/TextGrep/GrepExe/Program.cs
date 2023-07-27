using Grep.Common;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

/*
 * ■入力ファイルフォーマット
 * 
 * ファイル名:input.tsv
 * 形式:tsv
 * エンコード:UTF-8 BOMあり
 * ヘッダ:なし
 * GrepDirPath\t(GREPフォルダの絶対パス)
 * IsDebugLog\t(0/1=デバッグログ出力しない/出力する)
 * IsSubDir\t(0/1=サブフォルダを検索しない/検索する)
 * TargetFile\t(対象ファイル。ワイルドカード指定可。1種類のみ)
 * IsRegEx\t(0/1=正規表現を使わない/使う)
 * IsIgnoreCase\t(0/1=大文字小文字を区別する/区別しない)
 * IsOutputZeroCount\t(0/1=0件は出力しない/出力する)
 * Keywords\t(GREPキーワード1。複数個指定可)
 * …
 * Keywords\t(GREPキーワードn)
 * 
 * 
 * ■出力ファイルフォーマット
 * 
 * ファイル名:output.tsv
 * 形式:tsv
 * エンコード:UTF-8 BOMあり
 * ヘッダ:あり(#\tFilePath\tFileName\tExtension\tKeyword\tHitCount)
 * 
 * 
 * ■その他
 * 対象ファイルのエンコードは以下のみとする。
 * UTF-8, Shift-JIS
 * 
 */

namespace Grep
{
    internal class Program
    {
        private const string OUTPUT_FILENAME = @"output.tsv";

        private static string _input_file_path;
        private static MainParam _main_param;
        private static List<Result> _results;

        /// <summary>
        /// メイン
        /// </summary>
        /// <param name="args">入力ファイルの絶対パス(Ex."C:\tmp\input.tsv")</param>
        /// <returns>0:成功, 0以外:失敗</returns>
        static int Main(string[] args)
        {
            int ret = 0;

            try
            {
                Logger.Initialize();
                Logger.Info("■Start");

                run(args);
            }
            catch (Exception ex)
            {
                if (!Logger.IsInitSuccess)
                {
                    Logger.Initialize();
                }

                Logger.Error($"Error! msg={ex.Message}, stack={ex.StackTrace}");
                ret = 1;
            }
            finally
            {
                Logger.Info("■End");
                Logger.Dispose();
            }

            return ret;
        }

        /// <summary>
        /// 文字列から検索文字列がいくつあるかを返す
        /// </summary>
        /// <param name="target">文字列</param>
        /// <param name="is_ignorecase">true:大文字小文字を区別しない, false:区別する</param>
        /// <param name="strArray">検索文字列</param>
        /// <returns>結果</returns>
        private static int countOf(string target, bool is_ignorecase,  params string[] strArray)
        {
            var count = 0;
            var option = StringComparison.CurrentCulture;
            if (is_ignorecase)
                option = StringComparison.CurrentCultureIgnoreCase;

            foreach (var str in strArray)
            {
                var index = target.IndexOf(str, 0, option);
                while (index != -1)
                {
                    count++;
                    index = target.IndexOf(str, index + str.Length, option);
                }
            }

            return count;
        }

        /// <summary>
        /// 実行
        /// </summary>
        /// <param name="args"></param>
        /// <exception cref="ArgumentException"></exception>
        private static void run(string[] args)
        {
            Logger.Debug("run S");

            if (args.Length < 1)
            {
                Console.WriteLine(@"Grep.exe Usage");
                Console.WriteLine(@"Arg1:Full path of 'input.tsv' file.");
                Console.WriteLine(@"[FILE FORMAT of input.tsv ]");
                Console.WriteLine(@"* ファイル名:input.tsv");
                Console.WriteLine(@"* 形式:tsv");
                Console.WriteLine(@"* エンコード:UTF-8 with BOM");
                Console.WriteLine(@"* ヘッダ:なし");
                Console.WriteLine(@"* GrepDirPath\t(GREPフォルダの絶対パス)");
                Console.WriteLine(@"* IsDebugLog\t(0/1=デバッグログ出力しない/出力する)");
                Console.WriteLine(@"* IsSubDir\t(0/1=サブフォルダを検索しない/検索する)");
                Console.WriteLine(@"* TargetFile\t(対象ファイル。ワイルドカード指定可。1種類のみ)");
                Console.WriteLine(@"* IsRegEx\t(0/1=正規表現を使わない/使う)");
                Console.WriteLine(@"* IsIgnoreCase\t(0/1=大文字小文字を区別する/区別しない)");
                Console.WriteLine(@"* IsOutputZeroCount\t(0/1=0件は出力しない/出力する)");
                Console.WriteLine(@"* Keywords\t(GREPキーワード1。複数個指定可)");
                Console.WriteLine(@"* …");
                Console.WriteLine(@"* Keywords\t(GREPキーワードn)");
                return;
            }

            _input_file_path = args[0];

            Logger.Info($"args[0]={_input_file_path}");

            parseInputFile();

            grep();

            createOutputFile();

            Logger.Debug("run E");
        }

        /// <summary>
        /// GREPメイン
        /// </summary>
        private static void grep()
        {
            Logger.Debug("grep S");

            _results = new List<Result>();

            IEnumerable<FileInfo> fileList = 
                getFiles(_main_param.GrepDirPath, _main_param.TargetFile, _main_param.IsSubDir);

            foreach(var keyword in _main_param.Keywords)
            {
                var option = RegexOptions.IgnoreCase;
                if (!_main_param.IsIgnoreCase)
                    option = RegexOptions.None;
                Regex regex = new Regex(keyword, option);

                var queryMatchingFiles =
                    fileList
                        .Select(file =>
                        {
                            var contents = ReadContents(file.FullName);

                            if (!_main_param.IsRegEx)
                            {
                                return new
                                {
                                    filepath = file.FullName,
                                    hitcount = countOf(contents, _main_param.IsIgnoreCase, keyword)
                                };
                            }
                            else
                            {
                                var matches = regex.Matches(contents);

                                return new
                                {
                                    filepath = file.FullName,
                                    hitcount = matches.Cast<Match>().Select(match => match.Value).Count(),
                                };
                            }
                        });

                foreach (var v in queryMatchingFiles)
                {
                    _results.Add(new Result(v.filepath, keyword, v.hitcount));
                }
            }

            Logger.Debug("grep E");
        }

        /// <summary>
        /// GREP結果をファイル出力する
        /// </summary>
        private static void createOutputFile()
        {
            Logger.Debug("createOutputFile S");

            var path = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), OUTPUT_FILENAME);

            Logger.Info($"output.tsv:{path}");

            if (File.Exists(path))
            {
                File.Delete(path);
            }

            var num = 1;

            using (var sw = new StreamWriter(path, false, new UTF8Encoding(true)))
            {
                // Header
                sw.WriteLine("#\tFilePath\tFileName\tExtension\tKeyword\tHitCount");

                // Body
                foreach (var result in _results)
                {
                    if (!_main_param.IsOutputZeroCount && result.HitCount == 0)
                        continue;

                    sw.WriteLine($"{num}\t{result.FilePath}\t{result.FileName}\t{result.Extension}\t{result.Keyword}\t{result.HitCount}");
                    num++;
                }
            }

            Logger.Debug("createOutputFile E");
        }

        /// <summary>
        /// フォルダ配下を検索してファイル一覧を返す
        /// </summary>
        /// <param name="search_dir">フォルダパス</param>
        /// <param name="target_file">対象ファイル(ワイルドカード可)</param>
        /// <param name="is_sub_dir">true:サブフォルダも検索。false:フォルダ直下のみ検索</param>
        /// <returns></returns>
        /// <exception cref="DirectoryNotFoundException"></exception>
        /// <exception cref="ArgumentException"></exception>
        private static IEnumerable<FileInfo> getFiles(string search_dir, string target_file, bool is_sub_dir = true)
        {
            Logger.Debug("getFiles S");

            if (!Directory.Exists(search_dir))
                throw new DirectoryNotFoundException($"フォルダが存在しない。({search_dir})");

            if (target_file == "")
                throw new ArgumentException("対象ファイルが未指定。");

            var files = new List<FileInfo>();

            var option = SearchOption.AllDirectories;
            if (!is_sub_dir)
                option = SearchOption.TopDirectoryOnly;

            var ext = Path.GetExtension(target_file);

            foreach (string name in Directory.GetFiles(search_dir, target_file, option))
            {
                if (endsWith(name, ext))
                    files.Add(new FileInfo(name));
            }

            Logger.Debug("getFiles E");

            return files;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="src"></param>
        /// <param name="end"></param>
        /// <returns></returns>
        private static bool endsWith(string src, string end)
        {
            int endIndex = src.Length - end.Length;
            return src.Substring(endIndex) == end;
        }

        /// <summary>
        /// コマンドライン引数に指定されたファイルの解析
        /// </summary>
        private static void parseInputFile()
        {
            Logger.Debug("parseInputFile S");

            parseMainParam();

            if (!_main_param.IsDebugLog)
            {
                Logger.Debug("parseInputFile E1");
                Logger.Info("■End");
                Logger.Dispose();
            }

            Logger.Debug("parseInputFile E");
        }

        /// <summary>
        /// メインパラメータの解析とデータモデルの作成
        /// </summary>
        private static void parseMainParam()
        {
            Logger.Debug("parseMainParam S");

            _main_param = new MainParam();

            using (var stream = File.OpenText(_input_file_path))
            {
                string line;
                while ((line = stream.ReadLine()) != null)
                {
                    string[] columns = line.Split('\t');

                    if (columns[0] == "GrepDirPath")
                    {
                        _main_param.GrepDirPath = columns[1];
                    }
                    else if (columns[0] == "IsDebugLog")
                    {
                        _main_param.IsDebugLog = columns[1] == "0" ? false : true;
                    }
                    else if (columns[0] == "IsSubDir")
                    {
                        _main_param.IsSubDir = columns[1] == "0" ? false : true;
                    }
                    else if (columns[0] == "TargetFile")
                    {
                        _main_param.TargetFile = columns[1];
                    }
                    else if (columns[0] == "IsRegEx")
                    {
                        _main_param.IsRegEx = columns[1] == "0" ? false : true;
                    }
                    else if (columns[0] == "IsIgnoreCase")
                    {
                        _main_param.IsIgnoreCase = columns[1] == "0" ? false : true;
                    }
                    else if (columns[0] == "IsOutputZeroCount")
                    {
                        _main_param.IsOutputZeroCount = columns[1] == "0" ? false : true;
                    }
                    else if (columns[0] == "Keywords")
                    {
                        _main_param.Keywords.Add(columns[1]);
                    }
                }
            }

            Logger.Debug("parseMainParam E");
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="filepath"></param>
        /// <returns></returns>
        private static string ReadContents(string filepath)
        {
            byte[] bs = File.ReadAllBytes(filepath);

            Encoding enc = DetectEncodingFromBOM(bs);

            if (enc != Encoding.UTF8)
            {
                enc = Encoding.GetEncoding(932);
            }

            return File.ReadAllText(filepath, enc);
        }

        /// <summary>
        /// BOMを調べて、文字コードを判別する。
        /// </summary>
        /// <param name="bytes">文字コードを調べるデータ。</param>
        /// <returns>BOMが見つかった時は、対応するEncodingオブジェクト。
        /// 見つからなかった時は、null。</returns>
        private static Encoding DetectEncodingFromBOM(byte[] bytes)
        {
            if (bytes.Length < 2)
            {
                return null;
            }
            if ((bytes[0] == 0xfe) && (bytes[1] == 0xff))
            {
                //UTF-16 BE
                return new UnicodeEncoding(true, true);
            }
            if ((bytes[0] == 0xff) && (bytes[1] == 0xfe))
            {
                if ((4 <= bytes.Length) &&
                    (bytes[2] == 0x00) && (bytes[3] == 0x00))
                {
                    //UTF-32 LE
                    return new UTF32Encoding(false, true);
                }
                //UTF-16 LE
                return new UnicodeEncoding(false, true);
            }
            if (bytes.Length < 3)
            {
                return null;
            }
            if ((bytes[0] == 0xef) && (bytes[1] == 0xbb) && (bytes[2] == 0xbf))
            {
                //UTF-8
                return new UTF8Encoding(true, true);
            }
            if (bytes.Length < 4)
            {
                return null;
            }
            if ((bytes[0] == 0x00) && (bytes[1] == 0x00) &&
                (bytes[2] == 0xfe) && (bytes[3] == 0xff))
            {
                //UTF-32 BE
                return new UTF32Encoding(true, true);
            }

            return null;
        }
    }

    /// <summary>
    /// GREP結果
    /// </summary>
    internal class Result
    {
        /// <summary>
        /// ファイルパス
        /// </summary>
        public string FilePath;

        /// <summary>
        /// ファイル名
        /// </summary>
        public string FileName; 

        /// <summary>
        /// 拡張子
        /// </summary>
        public string Extension;

        /// <summary>
        /// GREPキーワード
        /// </summary>
        public string Keyword;

        /// <summary>
        /// ヒット数
        /// </summary>
        public int HitCount;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="filePath">ファイルパス</param>
        /// <param name="keyword">GREPキーワード</param>
        /// <param name="hitCount">ヒット数</param>
        public Result(string filePath, string keyword, int hitCount)
        {
            FilePath = filePath;
            FileName = Path.GetFileName(filePath);
            Extension = Path.GetExtension(filePath);
            Keyword = keyword;
            HitCount = hitCount;
        }
    }

    /// <summary>
    /// メインパラメータ
    /// </summary>
    internal class MainParam
    {
        /// <summary>
        /// GREPフォルダパス
        /// </summary>
        public string GrepDirPath;

        /// <summary>
        /// デバッグログ出力
        /// </summary>
        public bool IsDebugLog;

        /// <summary>
        /// サブフォルダも含む
        /// </summary>
        public bool IsSubDir;

        /// <summary>
        /// 対象ファイル
        /// </summary>
        public string TargetFile;

        /// <summary>
        /// 正規表現
        /// </summary>
        public bool IsRegEx;

        /// <summary>
        /// 大文字小文字区別
        /// </summary>
        public bool IsIgnoreCase;

        /// <summary>
        /// 0件も出力する
        /// </summary>
        public bool IsOutputZeroCount;

        /// <summary>
        /// GREPキーワード
        /// </summary>
        public List<string> Keywords;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public MainParam()
        {
            Keywords = new List<string>();
        }
    }

}
