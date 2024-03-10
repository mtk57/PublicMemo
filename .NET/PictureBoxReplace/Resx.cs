using System.Collections.Generic;

namespace PictureBoxReplace
{
    internal class Resx
    {
        public static List<string> GetPictureBoxNames (string resxPath)
        {
            var ret = new List<string>();

            var resx = Common.ReadXmlFile( resxPath );

            var datas = resx.Root.Elements( "data" );

            foreach ( var data in datas )
            {
                var name = data.Attribute( "name" ).Value;
                var type = data.Attribute( "type" ).Value;
                var mime = data.Attribute( "mimetype" ).Value;

                var nameSplit = name.Split( '.' );
                if (nameSplit.Length != 2 || nameSplit[1] != "Image") continue;

                if ( type != Common.ATTR_TYPE || mime != Common.ATTR_MIMETYPE ) continue;

                ret.Add( nameSplit[0] );
            }
            return ret;
        }

        public static string GetImageData ( string resxPath, string picName )
        {
            var ret = "";

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

                if ( nameSplit [ 0 ] != picName ) continue;

                ret = data.Element( "value" ).Value;

                break;
            }

            return ret.Replace( " ", "" ).Replace( "\n", "" ).Replace( "\r", "" );
        }


    }
}
