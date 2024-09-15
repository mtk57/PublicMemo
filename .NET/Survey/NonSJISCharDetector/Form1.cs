using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NonSJISCharDetector
{
    public partial class Form1 : Form
    {
        public Form1 ()
        {
            InitializeComponent();

#if DEBUG
            textBoxDirPath.Text = @"C:\_git\PublicMemo\.NET\Survey\NonSJISCharDetector\testdata";
            textBoxOutDirPath.Text = @"C:\_git\PublicMemo\.NET\Survey\NonSJISCharDetector\output";
#endif
        }

        private void buttonRefDir_Click ( object sender, EventArgs e )
        {
            this.textBoxDirPath.Text = ChooseDirPath();
        }

        private void buttonRefOutDir_Click ( object sender, EventArgs e )
        {
            this.textBoxOutDirPath.Text = ChooseDirPath();
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
            var folderPath = this.textBoxDirPath.Text;

            var extensions = this.textBoxExt.Text.Split( ',' );

            var outputPath = this.textBoxOutDirPath.Text + @"\" + DateTime.Now.ToString( "yyyyMMdd_HHmmss" ) + ".tsv";

            List<string> results = new List<string>();

            foreach ( var ext in extensions )
            {
                foreach ( var filePath in Directory.GetFiles( folderPath, $"*.{ext}", SearchOption.AllDirectories ) )
                {
                    FindInvalidShiftJISBytes( filePath, results );
                }
            }

            File.WriteAllLines( outputPath, results, Encoding.UTF8 );
        }

        private void FindInvalidShiftJISBytes ( string filePath, List<string> results )
        {
            List<(long offset, byte value)> invalidData = new List<(long, byte)>();

            using ( FileStream fs = new FileStream( filePath, FileMode.Open, FileAccess.Read ) )
            {
                byte [] buffer = new byte [ 1024 ];
                int bytesRead;
                long offset = 0;

                while ( ( bytesRead = fs.Read( buffer, 0, buffer.Length ) ) > 0 )
                {
                    for ( int i = 0; i < bytesRead; i++ )
                    {
                        byte b1 = buffer [ i ];
                        if ( IsShiftJISLeadByte( b1 ) )
                        {
                            if ( i + 1 < bytesRead )
                            {
                                byte b2 = buffer [ i + 1 ];
                                if ( !IsValidShiftJISCharacter( b1, b2 ) )
                                {
                                    invalidData.Add( (offset + i, b1) );
                                }
                                i++; // 2バイト文字なので、次のバイトをスキップ
                            }
                            else
                            {
                                // ファイルの最後で不完全な2バイト文字がある場合
                                invalidData.Add( (offset + i, b1) );
                            }
                        }
                        else if ( !IsValidShiftJISSingleByte( b1 ) )
                        {
                            invalidData.Add( (offset + i, b1) );
                        }
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
        }

        private bool IsShiftJISLeadByte ( byte b )
        {
            return ( b >= 0x81 && b <= 0x9F ) || ( b >= 0xE0 && b <= 0xFC );
        }

        private bool IsValidShiftJISSingleByte ( byte b )
        {
            return ( b >= 0x20 && b <= 0x7E ) || ( b >= 0xA1 && b <= 0xDF );
        }

        private bool IsValidShiftJISCharacter ( byte b1, byte b2 )
        {
            if ( b1 >= 0x81 && b1 <= 0x9F )
            {
                return ( b2 >= 0x40 && b2 <= 0x7E ) || ( b2 >= 0x80 && b2 <= 0xFC );
            }
            else if ( b1 >= 0xE0 && b1 <= 0xFC )
            {
                return ( b2 >= 0x40 && b2 <= 0x7E ) || ( b2 >= 0x80 && b2 <= 0xFC );
            }
            return false;
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
