using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using TextCompressor.Common;

namespace TextCompressor
{
    internal class Compressor
    {
        private string keyword;
        private string targetDirPath;
        private string outputDirPath;
        private List<string> targetExtensions;
        private bool noEncrypt;

        public Compressor(string kword, string inPath, string outPath, string ext, bool noEnc = false)
        {
            keyword = kword;
            targetDirPath = inPath;
            outputDirPath = outPath;
            targetExtensions = ext.Split(Const.EXT_DELI).ToList();
            noEncrypt = noEnc;

            Validate();
        }

        private void Validate()
        {
            if (string.IsNullOrEmpty(keyword))
            {
                throw new ArgumentNullException("Keyword is nothing!");
            }
            if (string.IsNullOrEmpty(targetDirPath))
            {
                throw new ArgumentNullException("Target Dir Path is nothing!");
            }
            if (string.IsNullOrEmpty(outputDirPath))
            {
                throw new ArgumentNullException("Output Dir Path is nothing!");
            }
            if (targetExtensions.Count == 0)
            {
                throw new ArgumentNullException("Target tExtensions is nothing!");
            }

            if (!Directory.Exists(targetDirPath))
            {
                throw new DirectoryNotFoundException("Target Dir Path is not exist!");
            }
            if (!Directory.Exists(outputDirPath))
            {
                throw new DirectoryNotFoundException("Output Dir Path is not exist!");
            }
        }

        public void Run()
        {
            // ファイル一覧を取得する
            var fileList = Utils.GetFileList(targetDirPath, targetExtensions);

            if (fileList.Count() == 0)
            {
                throw new FileNotFoundException("File not found!");
            }

            var sb = new StringBuilder();

            foreach(var file in fileList)
            {
                // 以下フォルダは除外
                if (file.Contains(@"\bin\") || file.Contains(@"\obj\"))
                {
                    continue;
                }

                // ファイルを読み込み結合する
                // ファイルパスは目印を付ける
                sb.AppendLine(string.Format(Const.FORMAT_PATH, file.Replace(targetDirPath, "")));
                sb.Append(Utils.ReadFile(file));
                //sb.AppendLine();
            }

            if (noEncrypt)
            {
                // 暗号化しないで保存する
                Utils.WriteFile(
                    Path.Combine(outputDirPath, Guid.NewGuid().ToString() + Const.EXT_CMP),
                    sb.ToString());
            }
            else
            {
                // 暗号化して保存する
                Utils.WriteFile(
                    Path.Combine(outputDirPath, Guid.NewGuid().ToString() + Const.EXT_CMP),
                    Utils.EncryptString(sb.ToString(), keyword));
            }
        }
    }
}
