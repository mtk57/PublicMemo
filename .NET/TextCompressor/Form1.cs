using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TextCompressor.Common;

namespace TextCompressor
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                Logger.Initialize();
            }
            catch (Exception ex)
            {
                OutException(ex);
            }
        }

        private void OutException(Exception ex)
        {
            var msg = $"Error! Messge={ex.Message}, Stack={ex.StackTrace}";

            if (Logger.IsInitSuccess)
            {
                Logger.Error(msg);
            }

            MessageBox.Show(msg);
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            try
            {
                Logger.Dispose();
            }
            catch (Exception ex)
            {
                OutException(ex);
            }
        }

        private void buttonCompInDirRef_Click(object sender, EventArgs e)
        {
            try
            {
                //FolderBrowserDialogクラスのインスタンスを作成
                var fbd = new FolderBrowserDialog();

                //上部に表示する説明テキストを指定する
                fbd.Description = "Select the target root directory.";
                //ルートフォルダを指定する
                //デフォルトでDesktop
                fbd.RootFolder = Environment.SpecialFolder.Desktop;
                //最初に選択するフォルダを指定する
                //RootFolder以下にあるフォルダである必要がある
                fbd.SelectedPath = @"C:\Windows";
                //ユーザーが新しいフォルダを作成できるようにする
                //デフォルトでTrue
                fbd.ShowNewFolderButton = true;

                //ダイアログを表示する
                if (fbd.ShowDialog(this) == DialogResult.OK)
                {
                    //選択されたフォルダを表示する
                    //Console.WriteLine(fbd.SelectedPath);
                    textBoxCompInDirPath.Text = fbd.SelectedPath;
                }
            }
            catch (Exception ex)
            {
                OutException(ex);
            }
        }

        private void buttonCompOutDirRef_Click(object sender, EventArgs e)
        {
            try
            {
                //FolderBrowserDialogクラスのインスタンスを作成
                var fbd = new FolderBrowserDialog();

                //上部に表示する説明テキストを指定する
                fbd.Description = "Select the directory to output the compressed file.";
                //ルートフォルダを指定する
                //デフォルトでDesktop
                fbd.RootFolder = Environment.SpecialFolder.Desktop;
                //最初に選択するフォルダを指定する
                //RootFolder以下にあるフォルダである必要がある
                fbd.SelectedPath = @"C:\Windows";
                //ユーザーが新しいフォルダを作成できるようにする
                //デフォルトでTrue
                fbd.ShowNewFolderButton = true;

                //ダイアログを表示する
                if (fbd.ShowDialog(this) == DialogResult.OK)
                {
                    //選択されたフォルダを表示する
                    //Console.WriteLine(fbd.SelectedPath);
                    textBoxCompOutDirPath.Text = fbd.SelectedPath;
                }
            }
            catch (Exception ex)
            {
                OutException(ex);
            }
        }

        private void buttonRunComp_Click(object sender, EventArgs e)
        {
            try
            {
                var c = new Compressor(textBoxKeyword.Text, textBoxCompInDirPath.Text, textBoxCompOutDirPath.Text, textBoxCompExt.Text);
                c.Run();

                MessageBox.Show("Success!");
            }
            catch (Exception ex)
            {
                OutException(ex);
            }
        }

        private void buttonDecompInFileRef_Click(object sender, EventArgs e)
        {
            try
            {
                //OpenFileDialogクラスのインスタンスを作成
                var ofd = new OpenFileDialog();

                //はじめのファイル名を指定する
                //はじめに「ファイル名」で表示される文字列を指定する
                //ofd.FileName = "default.html";
                //はじめに表示されるフォルダを指定する
                //指定しない（空の文字列）の時は、現在のディレクトリが表示される
                ofd.InitialDirectory = @"C:\";
                //[ファイルの種類]に表示される選択肢を指定する
                //指定しないとすべてのファイルが表示される
                ofd.Filter = "Compressファイル(*.cmp)|*.cmp";
                //[ファイルの種類]ではじめに選択されるものを指定する
                //2番目の「すべてのファイル」が選択されているようにする
                //ofd.FilterIndex = 2;
                //タイトルを設定する
                ofd.Title = "Select a compressed file.";
                //ダイアログボックスを閉じる前に現在のディレクトリを復元するようにする
                ofd.RestoreDirectory = true;
                //存在しないファイルの名前が指定されたとき警告を表示する
                //デフォルトでTrueなので指定する必要はない
                ofd.CheckFileExists = true;
                //存在しないパスが指定されたとき警告を表示する
                //デフォルトでTrueなので指定する必要はない
                ofd.CheckPathExists = true;

                //ダイアログを表示する
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    //OKボタンがクリックされたとき、選択されたファイル名を表示する
                    //Console.WriteLine(ofd.FileName);
                    textBoxDecompInFilePath.Text = ofd.FileName;
                }
            }
            catch (Exception ex)
            {
                OutException(ex);
            }
        }

        private void buttonDecompOutDirRef_Click(object sender, EventArgs e)
        {
            try
            {
                //FolderBrowserDialogクラスのインスタンスを作成
                var fbd = new FolderBrowserDialog();

                //上部に表示する説明テキストを指定する
                fbd.Description = "Select the root directory where you want to extract the compressed file.";
                //ルートフォルダを指定する
                //デフォルトでDesktop
                fbd.RootFolder = Environment.SpecialFolder.Desktop;
                //最初に選択するフォルダを指定する
                //RootFolder以下にあるフォルダである必要がある
                fbd.SelectedPath = @"C:\Windows";
                //ユーザーが新しいフォルダを作成できるようにする
                //デフォルトでTrue
                fbd.ShowNewFolderButton = true;

                //ダイアログを表示する
                if (fbd.ShowDialog(this) == DialogResult.OK)
                {
                    //選択されたフォルダを表示する
                    //Console.WriteLine(fbd.SelectedPath);
                    textBoxDecompOutDirPath.Text = fbd.SelectedPath;
                }
            }
            catch (Exception ex)
            {
                OutException(ex);
            }
        }

        private void buttonRunDecomp_Click(object sender, EventArgs e)
        {
            try
            {
                var dc = new Decompressor(textBoxKeyword.Text, textBoxDecompInFilePath.Text, textBoxDecompOutDirPath.Text);
                dc.Run();

                MessageBox.Show("Success!");
            }
            catch (Exception ex)
            {
                OutException(ex);
            }
        }

        private void buttonDefault_Click(object sender, EventArgs e)
        {
            textBoxKeyword.Text = "hoge";
            textBoxCompInDirPath.Text = @"C:\_git\Lab\C#\HttpClient2_Lib";
            textBoxCompExt.Text = "cs";
            textBoxCompOutDirPath.Text = @"C:\_tmp";
            textBoxDecompInFilePath.Text = @"C:\_tmp\test.cmp";
            textBoxDecompOutDirPath.Text = @"C:\_tmp";
        }
    }
}
