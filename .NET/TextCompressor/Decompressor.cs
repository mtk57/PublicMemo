using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using TextCompressor.Common;

namespace TextCompressor
{
    internal class Decompressor
    {
        private string keyword;
        private string targetFilePath;
        private string outputDirPath;
        private bool noEncrypt;

        public Decompressor(string kword, string inPath, string outPath, bool noEnc = false)
        {
            keyword = kword;
            targetFilePath = inPath;
            outputDirPath = outPath;
            noEncrypt = noEnc;

            Validate();
        }

        private void Validate()
        {
            if (string.IsNullOrEmpty(keyword))
            {
                throw new ArgumentNullException("Keyword is nothing!");
            }
            if (string.IsNullOrEmpty(targetFilePath))
            {
                throw new ArgumentNullException("Target File Path is nothing!");
            }
            if (string.IsNullOrEmpty(outputDirPath))
            {
                throw new ArgumentNullException("Output Dir Path is nothing!");
            }

            if (!File.Exists(targetFilePath))
            {
                throw new FileNotFoundException("Target File Path is not exist!");
            }
            if (!Directory.Exists(outputDirPath))
            {
                throw new DirectoryNotFoundException("Output Dir Path is not exist!");
            }
        }

        public void Run()
        {
            string cmpFile;

            if (noEncrypt)
            {
                // 圧縮ファイルを読み込む
                cmpFile = Utils.ReadFile(targetFilePath);
            }
            else
            {
                // 圧縮ファイルを読み込んで復号する
                cmpFile = Utils.DecryptString(
                                Utils.ReadFile(targetFilePath),
                                keyword);
            }

            StreamWriter sw = null;

            try
            {

                foreach (var line in Utils.SplitNewLine(cmpFile))
                {
                    var isFilePath = Regex.IsMatch(line, Const.FORMAT_PATH_REG);

                    if (!isFilePath)
                    {
                        // 出力モード

                        sw.WriteLine(line);
                        continue;
                    }
                    else
                    {
                        // ファイルパスを取得
                        var filePath = Path.Combine(
                                    outputDirPath,
                                    line.Replace(Const.MARK, "").Remove(0, 1)); // 先頭の\を削除

                        // 出力先のディレクトリを取得
                        var outDirPath = Path.GetDirectoryName(filePath);

                        if (!Directory.Exists(outDirPath))
                        {
                            try
                            {
                                var di = new DirectoryInfo(outDirPath);
                                di.Create();
                            }
                            catch
                            {
                                // TODO:ひとまず無視
                                continue;
                            }
                        }

                        if (File.Exists(filePath))
                        {
                            try
                            {
                                File.Delete(filePath);
                            }
                            catch
                            {
                                // TODO:ひとまず無視
                                continue;
                            }
                        }
                        if (sw != null)
                        {
                            sw.Close();
                            sw.Dispose();
                            sw = null;
                        }
                        sw = new StreamWriter(filePath, true, Encoding.UTF8);
                    }
                }


            }
            finally
            {
                if (sw != null)
                {
                    sw.Close();
                    sw.Dispose();
                    sw = null;
                }
            }
        }
    }
}
