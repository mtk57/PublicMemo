using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace NonSJISCharDetector
{
    public partial class Form1 : Form
    {
        public Form1 ()
        {
            InitializeComponent();

#if DEBUG
            //textBoxDirPath.Text = @"C:\_git\PublicMemo\.NET\Survey\NonSJISCharDetector\testdata";
            textBoxDirPath.Text = @"C:\_git\PublicMemo\.NET\Survey\NonSJISCharDetector\testdata2";
            textBoxOutDirPath.Text = @"C:\_git\PublicMemo\.NET\Survey\NonSJISCharDetector\output";
#endif
        }

        private void buttonRefDir_Click ( object sender, EventArgs e )
        {
            textBoxDirPath.Text = ChooseDirPath();
        }

        private void buttonRefOutDir_Click ( object sender, EventArgs e )
        {
            textBoxOutDirPath.Text = ChooseDirPath();
        }

        private void buttonRun_Click ( object sender, EventArgs e )
        {
            try
            {
                Run();

                MessageBox.Show( "Success" );
            }
            catch ( Exception ex )
            {
                MessageBox.Show( $"ERROR! {ex.Message}" );
            }
        }

        private void Run ()
        {
            var folderPath = textBoxDirPath.Text;

            var extensions = textBoxExt.Text
                            .Split( ',' )
                            .Select( ext => ext.Trim().ToLowerInvariant() )
                            .Where( ext => !string.IsNullOrEmpty( ext ) )
                            .ToArray();

            var outputPath = textBoxOutDirPath.Text + @"\" + DateTime.Now.ToString( "yyyyMMdd_HHmmss" ) + ".tsv";

            var results = new List<string>();

            foreach ( var ext in extensions )
            {
                foreach ( var filePath in Directory.GetFiles( folderPath, $"*.{ext}", SearchOption.AllDirectories ) )
                {
                    if ( IsSkipFileExtension( filePath, extensions ) ) continue;

                    FindAndReplaceInvalidShiftJISBytes( filePath, results );
                }
            }

            File.WriteAllLines( outputPath, results, Encoding.UTF8 );
        }

        private bool IsSkipFileExtension ( string filePath, string [] extensions )
        {
            return ! extensions.Contains( Path.GetExtension( filePath ).ToLower().Replace(".", "") );
        }

        // ------------------------------------------------------------------------
        // http://charset.7jp.net/sjis.html
        //
        // シフトJISの1バイトコード（半角文字）のエリア
        //   0x00～0x1F、0x7F：制御コード
        //   0x20～0x7E      ：ASCII
        //   0xA1～0xDF      ：半角カタカナ
        //
        // シフトJISの2バイトコード（全角文字）のエリア（JIS X 0208の漢字エリア）
        //　 上位1バイト　 0x81～0x9F、 0xE0～0xEF
        //　 下位1バイト　 0x40～0x7E、 0x80～0xFC
        //
        // ですが機種に依存しない観点より、以下エリアは使用しないのが無難です
        //   0x8540～ 0x889E
        //   0xEB40～ 0xEFFC
        //   0xF040～
        // ------------------------------------------------------------------------

        private void FindAndReplaceInvalidShiftJISBytes ( string filePath, List<string> results )
        {
            List<(long offset, byte value)> invalidData = new List<(long, byte)>();
            var fileModified = false;

            CreateBackupFileName( filePath );

            using ( FileStream fs = new FileStream( filePath, FileMode.Open, FileAccess.ReadWrite ) )
            {
                var buffer = new byte [ 1024 ];
                var bytesRead = 0;
                var offset = 0;

                while ( ( bytesRead = fs.Read( buffer, 0, buffer.Length ) ) > 0 )
                {
                    var bufferModified = false;

                    for ( var i = 0; i < bytesRead; i++ )
                    {
                        var b1 = buffer [ i ];

                        if ( IsShiftJISLeadByte( b1 ) )
                        {
                            if ( i + 1 < bytesRead )
                            {
                                byte b2 = buffer [ i + 1 ];
                                if ( !IsValidShiftJISCharacter( b1, b2 ) )
                                {
                                    invalidData.Add( (offset + i, b1) );
                                    ReplaceSpace( buffer, i );
                                    ReplaceSpace( buffer, i+1 );
                                    bufferModified = true;
                                    fileModified = true;
                                }
                                i++; // 2バイト文字なので、次のバイトをスキップ
                            }
                            else
                            {
                                // ファイルの最後で不完全な2バイト文字がある場合
                                invalidData.Add( (offset + i, b1) );
                                ReplaceSpace( buffer, i );
                                bufferModified = true;
                                fileModified = true;
                            }
                        }
                        else if ( !IsValidShiftJISSingleByte( b1 ) &&
                                  !IsExcludedControlCharacter( b1 ) )
                        {
                            invalidData.Add( (offset + i, b1) );
                            ReplaceSpace( buffer, i );
                            bufferModified = true;
                            fileModified = true;
                        }
                    }

                    if ( bufferModified )
                    {
                        fs.Position = offset;
                        fs.Write( buffer, 0, bytesRead );
                    }

                    offset += bytesRead;
                }
            }

            if ( invalidData.Any() )
            {
                foreach ( var (offset, value) in invalidData )
                {
                    results.Add( $"{filePath}\t{offset}\t0x{value:X2}" );
                }
            }

            if ( fileModified )
            {
                results.Add( $"{filePath}\tFile was modified" );
            }
        }

        // 全角文字の上位1バイトのみ
        private bool IsShiftJISLeadByte ( byte b )
        {
            return ( b >= 0x81 && b <= 0x9F ) ||
                   ( b >= 0xE0 && b <= 0xEF );
        }

        // 全角文字の上位1バイトと下位1バイト
        private bool IsValidShiftJISCharacter ( byte b1, byte b2 )
        {
            if ( b1 >= 0x81 && b1 <= 0x9F )
            {
                return ( b2 >= 0x40 && b2 <= 0x7E ) ||
                       ( b2 >= 0x80 && b2 <= 0xFC );
            }
            else if ( b1 >= 0xE0 && b1 <= 0xEF )
            {
                return ( b2 >= 0x40 && b2 <= 0x7E ) ||
                       ( b2 >= 0x80 && b2 <= 0xFC );
            }
            return false;
        }

        // 半角文字
        private bool IsValidShiftJISSingleByte ( byte b )
        {
            return ( b >= 0x20 && b <= 0x7E ) ||   // ASCII
                   ( b >= 0xA1 && b <= 0xDF );     // 半角カタカナ
        }

        // SPACE置換を除外する制御コード
        private bool IsExcludedControlCharacter ( byte b )
        {
            return b == 0x09 ||     // TAB
                   b == 0x0A ||     // LF
                   b == 0x0D;       // CR
        }

        private void CreateBackupFileName ( string filePath )
        {
            if ( !checkBoxCreateBackup.Checked ) return;

            var directory = Path.GetDirectoryName( filePath );
            var fileName = Path.GetFileNameWithoutExtension( filePath );
            var extension = Path.GetExtension( filePath );
            var timestamp = DateTime.Now.ToString( "yyyyMMddHHmmssfff" );

            var backupFilePath = Path.Combine( directory, $"{fileName}_{timestamp}{extension}" );

            File.Copy( filePath, backupFilePath, true );
        }

        private void ReplaceSpace ( byte [] array, int index )
        {
            if ( !checkBoxReplaceSpace.Checked ) return;

            array [ index ] = 0x20;
        }

        private string ChooseDirPath ()
        {
            var fbd = new FolderBrowserDialog();

            fbd.Description = "Choose directory";

            fbd.RootFolder = Environment.SpecialFolder.Desktop;

            fbd.SelectedPath = @"C:\Windows";

            fbd.ShowNewFolderButton = true;

            if ( fbd.ShowDialog( this ) == DialogResult.OK )
            {
                return fbd.SelectedPath;
            }

            return string.Empty;
        }

        private void textBoxDirPath_DragDrop ( object sender, DragEventArgs e )
        {
            var path = getDropData( e );

            if ( !Directory.Exists( path ) ) return;

            textBoxDirPath.Text = path;
        }

        private void textBoxDirPath_DragEnter ( object sender, DragEventArgs e )
        {
            startDragDrop( e );
        }

        private void textBoxOutDirPath_DragDrop ( object sender, DragEventArgs e )
        {
            var path = getDropData( e );

            if ( !Directory.Exists( path ) ) return;

            textBoxOutDirPath.Text = path;
        }

        private void textBoxOutDirPath_DragEnter ( object sender, DragEventArgs e )
        {
            startDragDrop( e );
        }

        private void startDragDrop ( DragEventArgs e )
        {
            if ( e.Data.GetDataPresent( DataFormats.FileDrop ) )
                e.Effect = DragDropEffects.Copy;
            else
                e.Effect = DragDropEffects.None;
        }

        private string getDropData ( DragEventArgs e )
        {
            var fileName = ( string [] ) e.Data.GetData( DataFormats.FileDrop, false );

            if ( fileName == null || fileName.Length == 0 || fileName.Length > 1 ) return string.Empty;

            return fileName [ 0 ];
        }
    }
}
