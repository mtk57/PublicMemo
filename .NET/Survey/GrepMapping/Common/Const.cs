namespace GrepMapping.Common
{
    internal class Const
    {
        public const string LOG = "GrepMapping.log";
        public const string CONFIG_JSON = "GrepMapping.json";

        public const string TYPE_SAKURA = "sakura";

        public const string CLM_APPLY_ORDER = "ClmApplyOrder";
        public const string CLM_ENABLE = "ClmEnable";
        public const string CLM_MAPPING_FILE_PATH = "ClmMappingFilePath";
        public const string CLM_SRC_START_ROW = "ClmSrcStartRow";
        public const string CLM_SRC_KEY = "ClmSrcKey";
        public const string CLM_SRC_COPY = "ClmSrcCopy";
        public const string CLM_DST_START_ROW = "ClmDstStartRow";
        public const string CLM_DST_KEY = "ClmDstKey";
        public const string CLM_DST_COPY = "ClmDstCopy";

        public const string CLM_GREP_RAW = "ClmGrepRaw";
        public const string CLM_GREP_FULL_PATH = "ClmGrepFullPath";
        public const string CLM_GREP_FILE_NAME = "ClmGrepFileName";
        public const string CLM_GREP_CONTENTS = "ClmGrepContents";
        public const string CLM_GREP_CUSTOM_ = "ClmGrepCustom_";

        public const string HEADER_APPLY_ORDER = "ApplyOrder";
        public const string HEADER_ENABLE = "Enable";
        public const string HEADER_MAPPING_FILE_PATH = "MappingFilePath";
        public const string HEADER_SRC_START_ROW = "SrcStartRow";
        public const string HEADER_SRC_KEY = "SrcKey";
        public const string HEADER_SRC_COPY = "SrcCopy";
        public const string HEADER_DST_START_ROW = "DstStartRow";
        public const string HEADER_DST_KEY = "DstKey";
        public const string HEADER_DST_COPY = "DstCopy";


    }

    internal class SakuraRegEx
    {
        public const string PTN_PATH = @"^([A-Z]:\\[^ ]+)";
        public const string PTN_CONTENTS = @"^.*\]:";
        public const string PTN_POS = @"\(\d+,\d+\)";
        public const string PTN_LTRIM = @"^[ \t]*";
    }

    internal enum Extension
    {
        UNKNOWN,
        VB
    }
}
