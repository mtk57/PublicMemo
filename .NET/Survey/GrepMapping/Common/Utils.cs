using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Windows.Forms;

namespace GrepMapping.Common
{
    internal static class Utils
    {
        public static void AddStringClm ( DataTable table, int count )
        {
            for ( int i = 0; i < count; i++ )
            {
                table.Columns.Add(i.ToString(), typeof(string));
            }
        }

        public static DataTable ConvertCsvToDataTable ( string path, char delimiter = ',', bool isHeader = true )
        {
            var ret = new DataTable();
            using ( var sr = new StreamReader( path ) )
            {
                var headers = sr.ReadLine().Split( delimiter );
                if ( isHeader )
                {
                    foreach ( var header in headers )
                    {
                        ret.Columns.Add( header );
                    }
                }
                while ( !sr.EndOfStream )
                {
                    var rows = sr.ReadLine().Split( delimiter );
                    var row = ret.NewRow();
                    for ( var i = 0; i < headers.Length; i++ )
                    {
                        row [ i ] = rows [ i ];
                    }
                    ret.Rows.Add( row );
                }
            }
            return ret;
        }

        public static string GetNowString ()
        {
            return DateTime.Now.ToString( "yyyyMMddHHmmssfff" );
        }

        public static void ExportDataTableToTsv ( DataTable dt, string filePath, bool isHeader = true )
        {
            using ( var writer = new StreamWriter( filePath, false, Encoding.UTF8 ) )
            {
                if ( isHeader )
                {
                    var header = string.Join( "\t", dt.Columns.Cast<DataColumn>().Select( col => col.ColumnName ) );
                    writer.WriteLine( header );
                }

                foreach ( DataRow row in dt.Rows )
                {
                    var line = string.Join( "\t", row.ItemArray.Select( field => field.ToString() ) );
                    writer.WriteLine( line );
                }
            }
        }

        public static Extension GetExtensionType ( string filename )
        {
            var ext = Path.GetExtension( filename ).ToLower();

            if ( ext == "bas" || ext == "frm" || ext == "cls" || ext == "ctl" || ext == "vb" )
            {
                return Extension.VB;
            }
            return Extension.UNKNOWN;
        }

        public static string GetDropData ( DragEventArgs e )
        {
            var data = ( string [] ) e.Data.GetData( DataFormats.FileDrop, false );

            if ( data == null || data.Length == 0 || data.Length > 1 ) return string.Empty;

            return data [ 0 ];
        }

        public static void StartDragDrop ( DragEventArgs e )
        {
            if ( e.Data.GetDataPresent( DataFormats.FileDrop ) )
                e.Effect = DragDropEffects.Copy;
            else
                e.Effect = DragDropEffects.None;
        }

        public static string OpenFileDialog ( string title, string filter )
        {
            var ofd = new OpenFileDialog
            {
                //FileName = "hoge.txt",
                InitialDirectory = @"C:\",
                //FilterIndex = 2,
                Filter = filter,
                Title = title,
                RestoreDirectory = true
            };

            if ( ofd.ShowDialog() == DialogResult.OK )
            {
                return ofd.FileName;
            }
            return string.Empty;
        }

        public static string FolderBrowserDialog ()
        {
            var fbd = new FolderBrowserDialog
            {
                Description = "フォルダを指定してください。",
                RootFolder = Environment.SpecialFolder.Desktop,
                SelectedPath = Path.GetDirectoryName( Application.ExecutablePath ),
                ShowNewFolderButton = true
            };

            if ( fbd.ShowDialog() == DialogResult.OK )
            {
                return fbd.SelectedPath;
            }
            return string.Empty;
        }

        public static DialogResult ShowOkCancelMessageBox ( string text = "実行します。よろしいですか？", string caption = "確認" )
        {
            return MessageBox.Show(
                text,
                caption,
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button2 );
        }

        public static string GetDotNetVersion()
        {
            return typeof(string).Assembly.GetName().Version.ToString();
        }

        public static string GetMyDir()
        {
            return Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        }

        public static T Deserialize<T>(string json)
        {
            T result;
            var serializer = new DataContractJsonSerializer(typeof(T));

            using (var ms = new MemoryStream(Encoding.UTF8.GetBytes(json)))
            {
                result = (T)serializer.ReadObject(ms);
            }
            return result;
        }

        public static string Serialize<T>(T obj)
        {
            using (var ms = new MemoryStream())
            {
                var serializer = new DataContractJsonSerializer(typeof(T));
                serializer.WriteObject(ms, obj);
                return Encoding.UTF8.GetString(ms.ToArray());
            }
        }

        public static string GetMyVersion()
        {
            var assembly = Assembly.GetExecutingAssembly().GetName();
            return assembly.Version.ToString();
        }
    }
}
