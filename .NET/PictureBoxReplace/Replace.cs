using System;
using System.Data;
using System.IO;

namespace PictureBoxReplace
{
    internal class Replace
    {
        private static DataModel _model = null;

        public static void Validate ( DataModel model )
        {
            _model = model;

            if ( string.IsNullOrEmpty( _model.ImageData ) )
            {
                throw new Exception( "ImageData is empty." );
            }

            var imageDate = _model.ImageData.Replace( " ", "" ).Replace( "\n", "" ).Replace( "\r", "" );

            if ( string.IsNullOrEmpty( imageDate ) )
            {
                throw new Exception( "ImageData is empty.(replace)" );
            }

            _model.ImageData = imageDate;

            if ( string.IsNullOrEmpty( _model.ReplaceImagePath ) )
            {
                throw new Exception( "ReplaceImagePath is empty." );
            }

            if ( !File.Exists( _model.ReplaceImagePath ) )
            {
                throw new Exception( $"ReplaceImagePath is not exist.\r\n({_model.ReplaceImagePath})" );
            }

            if ( string.IsNullOrEmpty( _model.TargetDirPath ) )
            {
                throw new Exception( "TargetDirPath is empty." );
            }

            if ( !Directory.Exists( _model.TargetDirPath ) )
            {
                throw new Exception( $"TargetDirPath is not exist.\r\n({_model.TargetDirPath})" );
            }

            if ( !Common.IsDir( _model.TargetDirPath ) )
            {
                throw new Exception( $"TargetDirPath is not directory.\r\n({_model.TargetDirPath})" );
            }
        }

        public static void Prepare ()
        {
            _model.ReplaceImageData = Common.ConvertImageToBase64String(_model.ReplaceImagePath);
        }

        public static DataTable Execute ()
        {
            var ret = CreateResultTable();

            foreach ( var resxPath in Common.GetResxFiles(_model.TargetDirPath) )
            {
                var backFile = "";

                if ( _model.IsBackup )
                {
                    backFile = Common.CreateBackupFile( resxPath );
                }

                var resx = Common.ReadXmlFile( resxPath );

                var datas = resx.Root.Elements( "data" );

                foreach ( var data in datas )
                {
                    var name = data.Attribute( "name" ).Value;
                    var type = data.Attribute( "type" ).Value;
                    var mime = data.Attribute( "mimetype" ).Value;

                    var nameSplit = name.Split( '.' );
                    if ( nameSplit.Length != 2 || nameSplit [ 1 ] != "Image" ) continue;

                    if ( type != Common.ATTR_TYPE || mime != Common.ATTR_MIMETYPE ) continue;

                    var target = data.Element( "value" ).Value;

                    target = target.Replace( " ", "" ).Replace( "\n", "" ).Replace( "\r", "" );

                    if ( target != _model.ImageData ) continue;

                    data.Element( "value" ).Value = _model.ReplaceImageData;

                    ret.Rows.Add( resxPath , nameSplit [0] );
                }

                if ( ret.Rows.Count > 0 )
                {
                    resx.Save( resxPath );
                }
                else
                {
                    if ( File.Exists( backFile ) )
                    {
                        try
                        {
                            File.Delete( backFile );
                        }
                        catch
                        {
                            // Do nothing
                        }
                    }
                }
            }// foreach

            return ret;
        }

        private static DataTable CreateResultTable ()
        {
            var ret = new DataTable ();

            ret.Columns.Add( Common.CLM_FILEPATH, typeof( string ) );
            ret.Columns.Add( Common.CLM_NAME, typeof( string ) );

            return ret;
        }
    }
}
