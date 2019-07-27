using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace FilteringGridTestApp
{
    public static class Utils
    {
        public const string FILTER_ALL = "すべてのファイル(*.*) | *.*";

        public const string FILTER_SQL = "SQLファイル(*.sql) | *.sql";

        private const string CONDITION_EQUAL = " {0} = '{1}' ";

        public enum Separeator
        {
            and,
            or
        }

        public static List<string> GetLayoutIdList ( DataTable table )
        {
            var list = new List<string>();

            foreach ( DataRow row in table.Rows )
            {
                if ( ( bool ) row [ LayoutInfo.NameIsDelete ] == false ) continue;

                var layoutId = row [ LayoutInfo.NameLayoutId ].ToString();

                list.Add( layoutId );
            }

            return list;
        }

        public static List<string> ReadTextFile (
            string filePath,
            Encoding encode = null )
        {
            var readData = new List<string>();

            Encoding encoding = Encoding.Default;

            if ( encode != null )
            {
                encoding = encode;
            }

            using ( var sr = new StreamReader( filePath, encoding ) )
            {
                while ( sr.Peek() != -1 )
                {
                    readData.Add( sr.ReadLine() );
                }
            }

            return readData;
        }

        public static void WriteTextFile (
            string filePath,
            List<string> writeData,
            bool isAppend = false,
            Encoding encode = null )
        {
            Encoding encoding = Encoding.Default;

            if ( encode != null )
            {
                encoding = encode;
            }

            using ( var sw = new StreamWriter( filePath, isAppend, encoding ) )
            {
                foreach ( var line in writeData )
                {
                    sw.WriteLine( line );
                }
            }
        }

        public static string GetExceptionMessage ( Exception ex ) => string.Format( $"{ex.Message}\n{ex.StackTrace}" );

        public static void ShowInfo ( string 
            message, string 
            caption = "Information" )
        {
            MessageBox.Show( message, caption, MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly );
        }

        public static DialogResult ShowWarning ( 
            string message, string 
            caption = "Warning" )
        {
            return MessageBox.Show( message, caption, MessageBoxButtons.OKCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly );
        }

        public static void ShowError ( 
            string message, 
            string caption = "Error" )
        {
            MessageBox.Show( message, caption, MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly );
        }

        public static string CreateFilterCondition ( 
            string clmName,
            string clmValue )
        {
            if ( string.IsNullOrEmpty( clmValue ) ) return string.Empty;

            return string.Format( CONDITION_EQUAL, clmName, clmValue );
        }

        public static string JoinCondition ( 
            Separeator separator, 
            params string [] conditions )
        {
            var list = new List<string>();

            foreach ( var condition in conditions )
            {
                if ( string.IsNullOrEmpty( condition ) ) continue;

                list.Add( condition );
            }

            return string.Join( separator.ToString(), list );
        }

        public static DataTable Distinct ( this DataTable table, params string [] clmNames )
        {
            return table.DefaultView.ToTable( true, clmNames );
        }

        public static string SelectFileDialog ( 
            string filter, 
            string caption = "", 
            string firstDir = @"C:\" )
        {
            var ofd = new OpenFileDialog();

            //ofd.FileName = "default.html";

            //ofd.InitialDirectory = firstDir;

            ofd.Filter = filter;

            ofd.Title = caption;

            ofd.RestoreDirectory = true;

            //ofd.CheckFileExists = true;

            //ofd.CheckPathExists = true;

            if ( ofd.ShowDialog() == DialogResult.OK )
            {
                return ofd.FileName;
            }

            return string.Empty;
        }

        public static string SaveFileDialog ( 
            string filter, 
            string caption = "", 
            string firstDir = @"C:\",
            string fileName = "")
        {
            var sfd = new SaveFileDialog();

            sfd.FileName = fileName;

            //sfd.InitialDirectory = firstDir;

            sfd.Filter = filter;

            //sfd.FilterIndex = 2;

            sfd.Title = caption;

            sfd.RestoreDirectory = true;

            if ( sfd.ShowDialog() == DialogResult.OK )
            {
                return sfd.FileName;
            }
            return string.Empty;
        }

        public static DataTable ConvertCsvToDataTable (
            string csvPath,
            char delimiter = ',',
            bool isHeader = true )
        {
            var dt = new DataTable();

            using ( var sr = new StreamReader( csvPath ) )
            {
                var line = sr.ReadLine().Split( delimiter );

                var columnCount = 0;

                foreach ( var clm in line )
                {
                    if ( isHeader )
                    {
                        dt.Columns.Add( clm );
                    }
                    else
                    {
                        // ヘッダ有無がfalseの場合、列名は数値とする(0～)
                        dt.Columns.Add( columnCount.ToString() );
                    }
                    columnCount++;
                }
            }
            using ( var sr = new StreamReader( csvPath ) )
            {
                var rowCnt = 0;

                while ( !sr.EndOfStream )
                {
                    var rows = sr.ReadLine().Split( delimiter );

                    if ( rowCnt == 0 && isHeader )
                    {
                        rowCnt++;
                        continue;
                    }

                    var dr = dt.NewRow();
                    for ( var i = 0; i < dt.Columns.Count; i++ )
                    {
                        dr [ i ] = rows [ i ];
                    }
                    dt.Rows.Add( dr );

                    rowCnt++;
                }
            }
            return dt;
        }
    }

    public class LayoutInfo
    {
        public const string ColumnPrefix = "Column";

        public string LayoutId { get; set; } = string.Empty;
        public string LayoutName { get; set; } = string.Empty;
        public string ServiceId { get; set; } = string.Empty;
        public string ServiceName { get; set; } = string.Empty;
        public string ProductId { get; set; } = string.Empty;
        public string ProductName { get; set; } = string.Empty;
        public bool IsDelete { get; set; } = false;

        public const string NameLayoutId = nameof( LayoutId );
        public const string NameLayoutName = nameof( LayoutName );
        public const string NameServiceId = nameof( ServiceId );
        public const string NameServiceName = nameof( ServiceName );
        public const string NameProductId = nameof( ProductId );
        public const string NameProductName = nameof( ProductName );
        public const string NameIsDelete = nameof( IsDelete );

        public const string ColumnLayoutId = ColumnPrefix + nameof( LayoutId );
        public const string ColumnLayoutName = ColumnPrefix + nameof( LayoutName );
        public const string ColumnServiceId = ColumnPrefix + nameof( ServiceId );
        public const string ColumnServiceName = ColumnPrefix + nameof( ServiceName );
        public const string ColumnProductId = ColumnPrefix + nameof( ProductId );
        public const string ColumnProductName = ColumnPrefix + nameof( ProductName );
        public const string ColumnIsDelete = ColumnPrefix + nameof( IsDelete );

        public LayoutInfo ()
        {
        }

        public LayoutInfo (
            string layoutId,
            string layoutName,
            string serviceId,
            string serviceName,
            string productId,
            string productName,
            bool isDelete )
        {
            LayoutId = layoutId;
            LayoutName = layoutName;
            ServiceId = serviceId;
            ServiceName = serviceName;
            ProductId = serviceId;
            ProductName = productName;
            IsDelete = isDelete;
        }
    }
}
