using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NonSJISCharDetector
{
    public partial class Form1 : Form
    {
        public Form1 ()
        {
            InitializeComponent();
        }

        private void Run ()
        {
            Console.WriteLine( "フォルダのパスを入力してください：" );
            string folderPath = Console.ReadLine();

            Console.WriteLine( "検索する拡張子を入力してください（例：.txt）：" );
            string extension = Console.ReadLine();

            Console.WriteLine( "結果を出力するTSVファイルのパスを入力してください：" );
            string outputPath = Console.ReadLine();

            List<string> results = new List<string>();

            foreach ( string filePath in Directory.GetFiles( folderPath, $"*{extension}", SearchOption.AllDirectories ) )
            {
                CheckFile( filePath, results );
            }

            File.WriteAllLines( outputPath, results, Encoding.UTF8 );
            Console.WriteLine( $"処理が完了しました。結果は {outputPath} に保存されました。" );
        }

        private void CheckFile ( string filePath, List<string> results )
        {
            byte [] fileBytes = File.ReadAllBytes( filePath );
            List<(int offset, byte value)> invalidData = new List<(int, byte)>();

            for ( int i = 0; i < fileBytes.Length; i++ )
            {
                byte b = fileBytes [ i ];
                if ( !IsValidShiftJISByte( b ) && !IsControlCharacter( b ) )
                {
                    invalidData.Add( (i, b) );
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

        private bool IsValidShiftJISByte ( byte b )
        {
            return ( b >= 0x20 && b <= 0x7E ) || ( b >= 0xA1 && b <= 0xDF ) || ( b >= 0x81 && b <= 0x9F ) || ( b >= 0xE0 && b <= 0xFC );
        }

        private bool IsControlCharacter ( byte b )
        {
            return b < 0x20 || b == 0x7F;
        }
    }


}
