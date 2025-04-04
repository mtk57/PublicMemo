using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SimpleFileSearch
{
    public partial class Form1: Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnBrowse_Click ( object sender, EventArgs e )
        {
            // フォルダ選択ダイアログを表示
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "検索するフォルダを選択してください";
                
                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    txtFolderPath.Text = folderDialog.SelectedPath;
                }
            }
        }

        private void btnSearch_Click ( object sender, EventArgs e )
        {
                        // バリデーションチェック
            if (string.IsNullOrWhiteSpace(txtKeyword.Text))
            {
                MessageBox.Show("検索キーワードを入力してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtKeyword.Focus();
                return;
            }

            if (string.IsNullOrWhiteSpace(txtFolderPath.Text))
            {
                MessageBox.Show("検索するフォルダを選択してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtFolderPath.Focus();
                return;
            }

            if (!Directory.Exists(txtFolderPath.Text))
            {
                MessageBox.Show("指定されたフォルダが存在しません。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtFolderPath.Focus();
                return;
            }

            // 以前の検索結果をクリア
            dataGridViewResults.Rows.Clear();

            try
            {
                Cursor = Cursors.WaitCursor;

                string searchPattern = txtKeyword.Text;
                bool useRegex = chkUseRegex.Checked;
                
                List<string> foundFiles = new List<string>();

                // ファイル検索
                if (useRegex)
                {
                    // 正規表現モード
                    try
                    {
                        Regex regex = new Regex(searchPattern, RegexOptions.IgnoreCase);
                        
                        foreach (string file in Directory.GetFiles(txtFolderPath.Text, "*.*", SearchOption.AllDirectories))
                        {
                            string fileName = Path.GetFileName(file);
                            if (regex.IsMatch(fileName))
                            {
                                foundFiles.Add(file);
                            }
                        }
                    }
                    catch (ArgumentException ex)
                    {
                        MessageBox.Show($"無効な正規表現です: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                else
                {
                    // ワイルドカードモード
                    try
                    {
                        // ワイルドカード検索にはWindows組み込みのワイルドカードサポートを使用
                        foundFiles.AddRange(Directory.GetFiles(txtFolderPath.Text, searchPattern, SearchOption.AllDirectories));
                    }
                    catch (ArgumentException ex)
                    {
                        MessageBox.Show($"無効な検索パターンです: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }

                // 結果を表示
                foreach (string file in foundFiles)
                {
                    dataGridViewResults.Rows.Add(file);
                }

                // 結果件数を表示
                this.Text = $"シンプルなファイル検索 - {foundFiles.Count} 件のファイルが見つかりました";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"エラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void dataGridViewResults_CellDoubleClick ( object sender, DataGridViewCellEventArgs e )
        {
            if (e.RowIndex >= 0 && dataGridViewResults.Rows[e.RowIndex].Cells[0].Value != null)
            {
                string filePath = dataGridViewResults.Rows[e.RowIndex].Cells[0].Value.ToString();
                
                if (File.Exists(filePath))
                {
                    try
                    {
                        // エクスプローラーでフォルダを開いてファイルを選択
                        Process.Start("explorer.exe", $"/select,\"{filePath}\"");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"フォルダを開けませんでした: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }
    }
}
