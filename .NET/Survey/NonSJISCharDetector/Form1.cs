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
            textBoxDirPath.Text = @"C:\_git\PublicMemo\.NET\Survey\NonSJISCharDetector\testdata";
            textBoxOutDirPath.Text = @"C:\_git\PublicMemo\.NET\Survey\NonSJISCharDetector\testdata";
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
        //   0x20～0x7E      ：ASCIIコード
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
        //
        // 1byte Check					
        //   00～08：SP置換
        //   09～0A：無視
        //   0B～0C：SP置換
        //   0D    ：無視
        //   0E～1F：SP置換
        //   20～7E：無視
        //   7F    ： SP置換
        //   80～8F：SP置換
        //   81～9F：上位byte
        //   A0～AF：SP置換
        //   A1～DF：無視
        //   E0～EF：上位byte
        //   F0～  ：SP置換
        //
        // 2byte Check					
        //   00～3F：SP置換 (1byte目も)
        //   40～7E：無視
        //   7F    ：SP置換 (1byte目も)
        //   80～FC：無視
        //   FD～  ：SP置換 (1byte目も)
        //
        // 機種依存
        //   8540～889E：SP置換 (1byte目も)
        //   EB40～EFFC：SP置換 (1byte目も)

        private void FindAndReplaceInvalidShiftJISBytes ( string filePath, List<string> results )
        {
            List<(long offset, byte value)> invalidData = new List<(long, byte)>();

            CreateBackupFileName( filePath );

            using ( var fs = new FileStream( filePath, FileMode.Open, FileAccess.ReadWrite ) )
            {
                var buffer = new byte [ 1024 ];
                var bytesRead = 0;
                var offset = 0;

                while ( ( bytesRead = fs.Read( buffer, 0, buffer.Length ) ) > 0 )
                {
                    var bufferModified = false;

                    for ( var i = 0; i < bytesRead; i++ )
                    {
                        // まずは1byteを判定する

                        var b1 = buffer [ i ];

                        if ( IsControlCode( b1 ) )
                        {
                            // 制御コード

                            if ( !IsIgnoreControlCode( b1 ) )
                            {
                                // 無視しない制御コード

                                invalidData.Add( (offset + i, b1) );
                                ReplaceSpace( buffer, i );
                                bufferModified = true;
                            }
                        }
                        else if ( IsSingleByteCode( b1 ) )
                        {
                            // 1byte文字 (ASCIIコード, 半角カタカナ)

                            // 無視する
                        }
                        else
                        {
                            // 上記以外

                            if ( IsHibyteCode( b1 ) )
                            {
                                // 2byte文字の上位byte

                                if ( i + 1 < bytesRead )
                                {
                                    byte b2 = buffer [ i + 1 ];

                                    if ( IsLowbyteCode( b2 ) )
                                    {
                                        // 2byte文字の下位byte

                                        if ( IsMachineDependentCode( b1, b2 ) )
                                        {
                                            // 機種依存コード

                                            invalidData.Add( (offset + i, b1) );
                                            ReplaceSpace( buffer, i );
                                            ReplaceSpace( buffer, i + 1 );
                                            bufferModified = true;
                                        }
                                    }
                                    i++;
                                }
                                else
                                {
                                    // ファイルの最後で不完全な2バイト文字がある場合

                                    invalidData.Add( (offset + i, b1) );
                                    ReplaceSpace( buffer, i );
                                    bufferModified = true;
                                }
                            }
                            else
                            {
                                // 2byte文字以外

                                invalidData.Add( (offset + i, b1) );
                                ReplaceSpace( buffer, i );
                                bufferModified = true;
                            }
                        }

                    }// for

                    if ( bufferModified )
                    {
                        fs.Position = offset;
                        fs.Write( buffer, 0, bytesRead );
                    }

                    offset += bytesRead;

                }// while

            }// stream

            if ( invalidData.Any() )
            {
                foreach ( var (offset, value) in invalidData )
                {
                    results.Add( $"{filePath}\t{offset}\t0x{value:X2}" );
                }
            }
        }

        // 制御コード (0x00～0x1F、0x7F)
        private bool IsControlCode ( byte b )
        {
            return ( b >= 0x00 && b <= 0x1F ) ||
                   ( b == 0x7F );
        }

        // 1バイト文字 (0x20～0x7E, 0xA1～0xDF)
        private bool IsSingleByteCode ( byte b )
        {
            return ( b >= 0x20 && b <= 0x7E ) ||
                   ( b >= 0xA1 && b <= 0xDF );
        }

        // 無視する制御コード
        private bool IsIgnoreControlCode ( byte b )
        {
            return b == 0x09 ||     // TAB
                   b == 0x0A ||     // LF
                   b == 0x0D;       // CR
        }

        // 2byte文字の上位byte
        private bool IsHibyteCode ( byte b )
        {
            return ( b >= 0x81 && b <= 0x9F ) ||
                   ( b >= 0xE0 && b <= 0xEF );
        }

        // 2byte文字の下位byte
        private bool IsLowbyteCode ( byte b )
        {
            return ( b >= 0x40 && b <= 0x7E ) ||
                   ( b >= 0x80 && b <= 0xFC );
        }

        // 機種依存コード
        private bool IsMachineDependentCode ( byte b1, byte b2 )
        {
            int code = ( b1 << 8 ) | b2;
            return ( code >= 0x8540 && code <= 0x889E ) ||
                   ( code >= 0xEB40 && code <= 0xEFFC );
        }

        private void CreateBackupFileName ( string filePath )
        {
            if ( !checkBoxCreateBackup.Checked ) return;

            var directory = Path.GetDirectoryName( filePath );
            var fileName = Path.GetFileNameWithoutExtension( filePath );
            var extension = Path.GetExtension( filePath );
            var timestamp = DateTime.Now.ToString( "yyyyMMddHHmmssfff" );

            var backupFilePath = Path.Combine( directory, $"{fileName}.{extension}.bak_{timestamp}" );

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
