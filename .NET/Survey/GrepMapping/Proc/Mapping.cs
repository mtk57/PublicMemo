using System.Collections.Generic;
using System.Data;
using System.Linq;
using GrepMapping.Common;
using GrepMapping.Config;

namespace GrepMapping.Proc
{
    internal class Mapping
    {
        private Configs _config;
        private MappingResultTable _result;
        private Dictionary<string, DataTable> _mappingFileDict;

        private Mapping ()
        {
        }

        public Mapping ( Configs config, MappingResultTable result)
        {
            _config = config;
            _result = result;
            _mappingFileDict = new Dictionary<string, DataTable> ();
        }

        public void Run ()
        {
            // 列を追加する
            _result.AddClms ( _config.GetMaxDstClm() );

            foreach ( MappingRuleInfo rule in _config.MappingRules.OrderBy( n => n.ApplyOrder ) )
            {
                if (!rule.IsEnable) continue;

                if ( !_mappingFileDict.ContainsKey( rule.MappingFilePath ) )
                {
                    // マッピングファイルを読み込んで辞書で保持しておく
                    _mappingFileDict.Add( rule.MappingFilePath, Utils.ConvertCsvToDataTable( rule.MappingFilePath, '\t' ) );
                }

                
                var mappingFile = _mappingFileDict [ rule.MappingFilePath ];

                for ( var i = 0; i < mappingFile.Rows.Count; i++ )
                {
                    // マッピングファイルからキーと値を取得
                    // ※2を引いている理由：ルールのStartRowは1始まり かつ ヘッダ部を含むが、DataTableは0始まり かつ ヘッダ部含まないため。
                    var srcKey = mappingFile.Rows [ i + rule.SrcStartRow - 2 ] [ rule.SrcKey - 1 ].ToString();
                    var srcVal = mappingFile.Rows [ i + rule.SrcStartRow - 2 ] [ rule.SrcCopy - 1 ].ToString();

                    // Grep結果からキーと一致する行を見つける
                    for ( var j = 0; j < _result.Rows.Count; j++ )
                    {
                        var dstKey = _result.Rows [ j + rule.DstStartRow - 1 ] [ rule.DstKey - 1 ].ToString();
                        var dstVal = _result.Rows [ j + rule.DstStartRow - 1 ] [ rule.DstCopy - 1 ].ToString();

                        if ( dstKey != srcKey ) continue;

                        // 発見

                        // 転記先が空?
                        if ( dstVal == "" )
                        {
                            // 転記元をコピー
                            _result.Rows [ j + rule.DstStartRow - 1 ] [ rule.DstCopy - 1 ] = srcVal;
                        }
                        else
                        {
                            if ( !rule.IsMultiLine ) continue;

                            // 複数行コピー

                            // 転記元と同じ?
                            if ( srcVal == dstVal )
                            {
                                // 転記しない
                            }
                            else
                            {
                                // 発見行を最終行にコピーしてから転記
                                var copiedRow = _result.Rows [ j + rule.DstStartRow - 1 ].ItemArray.Clone() as object [];
                                _result.Rows.Add( copiedRow );

                                _result.Rows [ _result.Rows.Count - 1 ] [ rule.DstCopy - 1 ] = srcVal;
                            }
                        }
                    }
                }
            }
        }
    }


}
