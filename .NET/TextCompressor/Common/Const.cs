namespace TextCompressor.Common
{
    internal class Const
    {
        public const string EXT_CMP = @".cmp";

        public const string MARK = "●";

        public static readonly string FORMAT_PATH = MARK + "{0}" + MARK;

        public static readonly string FORMAT_PATH_REG = "^" + MARK + @".*" + MARK + "$";

        public const char EXT_DELI = '|';
    }
}
