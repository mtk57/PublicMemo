using System;
using System.Linq;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Serialization;
using GrepMapping.Common;

namespace GrepMapping.Config
{
    internal static class ConfigUtil
    {
        public static string ConfigPath
        {
            get
            {
                return Path.Combine(Utils.GetMyDir(), Const.CONFIG_JSON);
            }
        }

        public static Configs ReadSettings()
        {
            if (!File.Exists(ConfigPath))
            {
                var msg = $"{Const.CONFIG_JSON} is not exist! ({ConfigPath})";
                throw new FileNotFoundException(msg);
            }
            var json = File.ReadAllText(ConfigPath);

            return Utils.Deserialize<Configs>(json);
        }

        public static void WriteSetting(Configs settings)
        {
            var json = Utils.Serialize(settings);

            File.WriteAllText(ConfigPath, json);
        }

        public static void Validate ( Configs settings )
        {
            // TODO
        }
    }

    [DataContract]
    public class Configs
    {
        [DataMember( Name = "grep_result_file_path" )]
        public string GrepResultFilePath { get; set; }

        [DataMember( Name = "mapping_result_dir_path" )]
        public string MappingResultDirPath { get; set; }

        [DataMember( Name = "grep_result_editing" )]
        public GrepResultEditingInfo GrepResultEditing { get; set; }

        [DataMember( Name = "mapping_rules" )]
        public List<MappingRuleInfo> MappingRules { get; set; }

        public int GetMaxDstClm ()
        {
            return MappingRules.OrderByDescending( x => x.GetMaxDstClm() ).FirstOrDefault().GetMaxDstClm();
        }

        public override string ToString()
        {
            return $"GrepResultFilePath=<{GrepResultFilePath}>, " +
                   $"MappingResultDirPath=<{MappingResultDirPath}>, " +
                   $"GrepResultEditing=<{GrepResultEditing}>, " +
                   $"MappingRules=<{MappingRules}>";
        }
    }

    [DataContract]
    public class GrepResultEditingInfo
    {
        [DataMember( Name = "grep_result_type" )]
        public string GrepResultType { get; set; }

        [DataMember( Name = "is_contents_line_comment_row_delete" )]
        public Boolean IsContentsLineCommentRowDelete { get; set; }

        [DataMember( Name = "replace_words" )]
        public List<ReplaceWordInfo> ReplaceWords { get; set; }

        public override string ToString ()
        {
            return $"GrepResultType=<{GrepResultType}>, " +
                   $"IsContentsLineCommentRowDelete=<{IsContentsLineCommentRowDelete}>, " +
                   $"ReplaceWords=<{ReplaceWords}>";
        }
    }

    [DataContract]
    public class ReplaceWordInfo
    {
        [DataMember( Name = "contents_replace_before_word" )]
        public string ContentsReplaceBeforeWord { get; set; }

        [DataMember( Name = "contents_replace_after_word" )]
        public string ContentsReplaceAfterWord { get; set; }

        [DataMember( Name = "is_regex" )]
        public Boolean IsRegEx { get; set; }

        public override string ToString ()
        {
            return $"ContentsReplaceBeforeWord=<{ContentsReplaceBeforeWord}>, " +
                   $"ContentsReplaceAfterWord=<{ContentsReplaceAfterWord}>, " +
                   $"IsRegEx=<{IsRegEx}>";
        }
    }

    [DataContract]
    public class MappingRuleInfo
    {
        [DataMember( Name = "apply_order" )]
        public int ApplyOrder { get; set; }

        [DataMember( Name = "is_enable" )]
        public Boolean IsEnable { get; set; }

        [DataMember( Name = "mapping_file_path" )]
        public string MappingFilePath { get; set; }

        [DataMember( Name = "src_start_row" )]
        public int SrcStartRow { get; set; }

        [DataMember( Name = "src_key" )]
        public int SrcKey { get; set; }

        [DataMember( Name = "src_copy" )]
        public int SrcCopy { get; set; }

        [DataMember( Name = "dst_start_row" )]
        public int DstStartRow { get; set; }

        [DataMember( Name = "dst_key" )]
        public int DstKey { get; set; }

        [DataMember( Name = "dst_copy" )]
        public int DstCopy { get; set; }

        public int GetMaxDstClm ()
        {
            if ( DstCopy >= DstKey ) return DstCopy;
            else return DstKey;
        }

        public override string ToString ()
        {
            return $"ApplyOrder=<{ApplyOrder}>, " +
                   $"IsEnable=<{IsEnable}>, " +
                   $"MappingFilePath=<{MappingFilePath}>, " +
                   $"SrcStartRow=<{SrcStartRow}>, " +
                   $"SrcKey=<{SrcKey}>, " +
                   $"SrcCopy=<{SrcCopy}>, " +
                   $"DstStartRow=<{DstStartRow}>, " +
                   $"DstKey=<{DstKey}>, " +
                   $"DstCopy=<{DstCopy}>";
        }
    }
}
