using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using GrepMapping.Config;
using GrepMapping.Common;

namespace GrepMapping.Proc
{
    internal class Output
    {
        private Configs _config;
        private MappingResultTable _result;

        private Output ()
        {
        }

        public Output ( Configs config, MappingResultTable result )
        {
            _config = config;
            _result = result;
        }

        public void Run ()
        {
            var path = Path.Combine( _config.MappingResultDirPath, Utils.GetNowString() + ".tsv" );

            Utils.ExportDataTableToTsv( _result, path);
        }
    }
}
