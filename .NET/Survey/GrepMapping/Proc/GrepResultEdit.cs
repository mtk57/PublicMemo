using System.IO;
using System.Text.RegularExpressions;
using GrepMapping.Common;
using GrepMapping.Config;

namespace GrepMapping.Proc
{
    internal class GrepResultEdit
    {
        private Configs _config;
        private MappingResultTable _result;

        private GrepResultEdit ()
        {
        }

        public GrepResultEdit ( Configs config, MappingResultTable result )
        {
            _config = config;

            if ( result == null )
            {
                result = new MappingResultTable ();
            }
            _result = result;
        }

        public void Run ()
        {
            Validate();

            UpdateTable();
        }

        private void Validate ()
        {
            if ( !File.Exists( _config.GrepResultFilePath ) )
            {
                throw new FileNotFoundException(
                    $"Grep result file is not exist! <{_config.GrepResultFilePath}>" );
            }
        }

        private void UpdateTable()
        {
            var raw = string.Empty;
            var fullpath = string.Empty;
            var filename = string.Empty;
            var contents = string.Empty;

            var lines = File.ReadAllLines( _config.GrepResultFilePath );
            foreach ( var line in lines )
            {
                var dr = _result.NewRow();

                raw = line.Replace("\t", " ");
                
                fullpath = GetFullPath( raw );
                if ( string.IsNullOrEmpty(fullpath ) ) continue;

                filename = GetFileName( fullpath );
                if ( string.IsNullOrEmpty(filename) ) continue;

                contents = GetContents( raw );
                if ( string.IsNullOrEmpty( contents ) ) continue;

                contents = IgnoreLine( filename, contents );
                if ( contents == null ) continue;

                dr [ Const.CLM_GREP_RAW ] = raw;
                dr [ Const.CLM_GREP_FULL_PATH ] = fullpath;
                dr [ Const.CLM_GREP_FILE_NAME ] = filename;
                dr [ Const.CLM_GREP_CONTENTS ] = contents;

                _result.Rows.Add( dr );
            }
        }

        private string GetFullPath ( string line )
        {
            string ret = null;

            if ( _config.GrepResultEditing.GrepResultType == Const.TYPE_SAKURA )
            {
                Match match = Regex.Match( line, SakuraRegEx.PTN_PATH );
                if ( match.Success )
                {
                    ret = Regex.Replace( match.Groups [ 0 ].Value, SakuraRegEx.PTN_POS, "");
                }
            }
            return ret;
        }

        private string GetFileName ( string fullpath )
        {
            return Path.GetFileName( fullpath );
        }

        private string GetContents ( string line )
        {
            string ret = null;

            if ( _config.GrepResultEditing.GrepResultType == Const.TYPE_SAKURA )
            {
                ret = Regex.Replace( line, SakuraRegEx.PTN_CONTENTS, "" );
                ret = Regex.Replace( ret, SakuraRegEx.PTN_LTRIM, "" );
            }
            return ret;
        }

        private string IgnoreLine ( string filename, string contents )
        {
            string ret = contents;

            if ( _config.GrepResultEditing.IsContentsLineCommentRowDelete)
            {
                if ( Utils.GetExtensionType( filename ) == Extension.VB && contents.StartsWith("'") )
                {
                    ret = null;
                }
            }
            return ret;
        }
    }
}
