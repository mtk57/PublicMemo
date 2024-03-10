using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Linq;

namespace PictureBoxReplace
{
    internal class Common
    {
        public const string ATTR_TYPE = "System.Drawing.Bitmap, System.Drawing";
        public const string ATTR_MIMETYPE = "application/x-microsoft.net.object.bytearray.base64";
        public const string CLM_FILEPATH = "FilePath";
        public const string CLM_NAME = "Name";

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

        public static string OpenFileDialog (string title, string filter)
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

        public static string ConvertImageToBase64String (string imagePath)
        {
            var ret = "";

            var picBox = new PictureBox
            {
                Image = Image.FromFile( imagePath )
            };

            var ext = Path.GetExtension( imagePath ).ToLower();

            using ( var ms = new MemoryStream() )
            {
                if ( ext == "png" )
                    picBox.Image.Save( ms, ImageFormat.Png );
                else if ( ext == "jpg" || ext == "jpeg" )
                    picBox.Image.Save( ms, ImageFormat.Jpeg );
                else
                    picBox.Image.Save( ms, ImageFormat.Bmp );

                ret = Convert.ToBase64String( ms.ToArray() );
            }

            if ( picBox.Image != null )
            {
                picBox.Image.Dispose();
                picBox.Image = null;
            }

            return ret;
        }

        public static Image ConvertBase64StringToImage ( string base64string )
        {
            Image ret = null;

            try
            {
                using ( var ms = new MemoryStream( Convert.FromBase64String( base64string ) ) )
                {
                    ret = Image.FromStream( ms );
                }
            }
            catch
            {
                ret?.Dispose();
                ret = null;
            }

            return ret;
        }

        public static XDocument ReadXmlFile ( string xmlPath )
        {
            return XDocument.Load( xmlPath );
        }

        public static IEnumerable<string> GetResxFiles ( string dirPath )
        {
            return Directory.EnumerateFiles( dirPath, "*.resx", SearchOption.AllDirectories )
                   .Where( f => !Path.GetFileName( f )
                   .Equals( "Resources.resx", StringComparison.OrdinalIgnoreCase ));
        }

        public static bool IsDir ( string path )
        {
            return File.GetAttributes( path ).HasFlag( FileAttributes.Directory );
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

        public static string CreateBackupFile ( string srcPath )
        {
            var dstPath = GetBackupFilePath( srcPath );
            int backupNumber = 1;

            while ( File.Exists( dstPath ) )
            {
                dstPath = GetBackupFilePath( srcPath, backupNumber );
                backupNumber++;
            }

            File.Copy( srcPath, dstPath );

            return dstPath;
        }

        private static string GetBackupFilePath ( string srcPath, int num = 0 )
        {
            var backupFileName = @".bak";

            if ( num > 0 )
            {
                backupFileName += num;
            }

            return srcPath + backupFileName;
        }
    }
}
