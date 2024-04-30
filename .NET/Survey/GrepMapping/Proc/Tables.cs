using System;
using System.Data;
using System.IO;
using GrepMapping.Common;
using GrepMapping.Config;

namespace GrepMapping.Proc
{
    internal class Tables
    {
        // Dummy
    }

    internal class MappingResultTable : DataTable
    {
        public MappingResultTable ()
        {
            CreateClms();
        }

        public void AddClms ( int count )
        {
            Utils.AddStringClm( this, count );
        }

        private void CreateClms ()
        {
            Columns.Add( Const.CLM_GREP_RAW, typeof( string ) );
            Columns.Add( Const.CLM_GREP_FULL_PATH, typeof( string ) );
            Columns.Add( Const.CLM_GREP_FILE_NAME, typeof( string ) );
            Columns.Add( Const.CLM_GREP_CONTENTS, typeof( string ) );
        }
    }

    internal class MappingDataTable : DataTable
    {
        private Configs _config = null;

        public MappingDataTable ()
        {
            CreateClms();
        }

        public MappingDataTable ( Configs config )
        {
            _config = config;

            CreateClms();

            UpdateFromConfig();
        }

        public bool Validate ()
        {
            int result_int;
            bool result_bool;

            foreach ( DataRow row in Rows )
            {
                if ( !bool.TryParse( row [ Const.CLM_ENABLE ].ToString(), out result_bool ) )
                {
                    throw new Exception( $"{Const.CLM_ENABLE} is bad value!!" );
                }

                if (!int.TryParse( row [ Const.CLM_APPLY_ORDER ].ToString(), out result_int ) || result_int < 1 )
                {
                    throw new Exception($"{Const.CLM_APPLY_ORDER} is bad value!!");
                }

                if ( row [ Const.CLM_MAPPING_FILE_PATH ].ToString() == "" )
                {
                    throw new Exception( $"{Const.CLM_MAPPING_FILE_PATH} is bad value!!" );
                }

                if (!File.Exists(row [ Const.CLM_MAPPING_FILE_PATH ].ToString()) )
                {
                    throw new Exception( $"{Const.CLM_MAPPING_FILE_PATH} is not exist!!" );
                }

                if ( !int.TryParse( row [ Const.CLM_SRC_START_ROW ].ToString(), out result_int ) || result_int < 1 )
                {
                    throw new Exception( $"{Const.CLM_SRC_START_ROW} is bad value!!" );
                }

                if ( !int.TryParse( row [ Const.CLM_SRC_KEY ].ToString(), out result_int ) || result_int < 1 )
                {
                    throw new Exception( $"{Const.CLM_SRC_KEY} is bad value!!" );
                }

                if ( !int.TryParse( row [ Const.CLM_SRC_COPY ].ToString(), out result_int ) || result_int < 1 )
                {
                    throw new Exception( $"{Const.CLM_SRC_COPY} is bad value!!" );
                }

                if ( !int.TryParse( row [ Const.CLM_DST_START_ROW ].ToString(), out result_int ) || result_int < 1 )
                {
                    throw new Exception( $"{Const.CLM_DST_START_ROW} is bad value!!" );
                }

                if ( !int.TryParse( row [ Const.CLM_DST_KEY ].ToString(), out result_int ) || result_int < 1 )
                {
                    throw new Exception( $"{Const.CLM_DST_KEY} is bad value!!" );
                }

                if ( !int.TryParse( row [ Const.CLM_DST_COPY ].ToString(), out result_int ) || result_int < 1 )
                {
                    throw new Exception( $"{Const.CLM_DST_COPY} is bad value!!" );
                }
            }

            return true;
        }

        private void CreateClms ()
        {
            Columns.Add( Const.CLM_APPLY_ORDER, typeof( int ) );
            Columns.Add( Const.CLM_ENABLE, typeof( bool ) );
            Columns.Add( Const.CLM_MAPPING_FILE_PATH, typeof( string ) );
            Columns.Add( Const.CLM_SRC_START_ROW, typeof( int ) );
            Columns.Add( Const.CLM_SRC_KEY, typeof( int ) );
            Columns.Add( Const.CLM_SRC_COPY, typeof( int ) );
            Columns.Add( Const.CLM_DST_START_ROW, typeof( int ) );
            Columns.Add( Const.CLM_DST_KEY, typeof( int ) );
            Columns.Add( Const.CLM_DST_COPY, typeof( int ) );
        }

        private void UpdateFromConfig ()
        {
            foreach ( MappingRuleInfo rule in _config.MappingRules )
            {
                var dr = NewRow();

                dr [ Const.CLM_APPLY_ORDER ] = rule.ApplyOrder;
                dr [ Const.CLM_ENABLE ] = rule.IsEnable;
                dr [ Const.CLM_MAPPING_FILE_PATH ] = rule.MappingFilePath;
                dr [ Const.CLM_SRC_START_ROW ] = rule.SrcStartRow;
                dr [ Const.CLM_SRC_KEY ] = rule.SrcKey;
                dr [ Const.CLM_SRC_COPY ] = rule.SrcCopy;
                dr [ Const.CLM_DST_START_ROW ] = rule.DstStartRow;
                dr [ Const.CLM_DST_KEY ] = rule.DstKey;
                dr [ Const.CLM_DST_COPY ] = rule.DstCopy;

                Rows.Add( dr );
            }
        }
    }
}
