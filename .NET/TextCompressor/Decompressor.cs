using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TextCompressor
{
    internal class Decompressor
    {
        private string keyword;
        private string targetFilePath;
        private string outputDirPath;

        public Decompressor(string kword, string inPath, string outPath)
        {
            keyword = kword;
            targetFilePath = inPath;
            outputDirPath = outPath;
        }

        public void Run()
        {

        }
    }
}
