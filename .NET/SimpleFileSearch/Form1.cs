using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace SimpleFileSearch
{
    public partial class Form1: Form
    {
        private const int MaxHistoryItems = 20;
        private const string SettingsFileName = "SimpleFileSearch.json";

        public Form1()
        {
            InitializeComponent();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            // 設定ファイルからデータを読み込む
            LoadSettings();
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 設定ファイルにデータを保存
            SaveSettings();
        }

        private void btnBrowse_Click ( object sender, EventArgs e )
        {
            // フォルダ選択ダイアログを表示
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "検索するフォルダを選択してください";
                
                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    cmbFolderPath.Text = folderDialog.SelectedPath;
                }
            }
        }

        private void btnSearch_Click ( object sender, EventArgs e )
        {
            // バリデーションチェック
            if (string.IsNullOrWhiteSpace(cmbKeyword.Text))
            {
                MessageBox.Show("検索キーワードを入力してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                cmbKeyword.Focus();
                return;
            }

            if (string.IsNullOrWhiteSpace(cmbFolderPath.Text))
            {
                MessageBox.Show("検索するフォルダを選択してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                cmbFolderPath.Focus();
                return;
            }

            if (!Directory.Exists(cmbFolderPath.Text))
            {
                MessageBox.Show("指定されたフォルダが存在しません。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                cmbFolderPath.Focus();
                return;
            }

            // 履歴を更新（キーワード）
            UpdateComboBoxHistory(cmbKeyword, cmbKeyword.Text);

            // 履歴を更新（フォルダパス）
            UpdateComboBoxHistory(cmbFolderPath, cmbFolderPath.Text);

            // 以前の検索結果をクリア
            dataGridViewResults.Rows.Clear();

            try
            {
                Cursor = Cursors.WaitCursor;

                string searchPattern = cmbKeyword.Text;
                bool useRegex = chkUseRegex.Checked;
                
                List<string> foundFiles = new List<string>();

                // ファイル検索
                if (useRegex)
                {
                    // 正規表現モード
                    try
                    {
                        Regex regex = new Regex(searchPattern, RegexOptions.IgnoreCase);
                        
                        foreach (string file in Directory.GetFiles(cmbFolderPath.Text, "*.*", SearchOption.AllDirectories))
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
                        foundFiles.AddRange(Directory.GetFiles(cmbFolderPath.Text, searchPattern, SearchOption.AllDirectories));
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
                        // エクスプローラーでフォルダを開いてファイルを選択（ツリーを展開）
                        OpenFolderAndSelectFile(filePath);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"フォルダを開けませんでした: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        #region コンボボックス履歴の管理

        private void UpdateComboBoxHistory(ComboBox comboBox, string item)
        {
            // 既に同じ項目が存在する場合は削除（重複を避けるため）
            if (comboBox.Items.Contains(item))
            {
                comboBox.Items.Remove(item);
            }

            // 最大数に達している場合は最も古い項目を削除
            if (comboBox.Items.Count >= MaxHistoryItems)
            {
                comboBox.Items.RemoveAt(comboBox.Items.Count - 1);
            }

            // 新しい項目を先頭に追加
            comboBox.Items.Insert(0, item);
            comboBox.Text = item;
        }

        #endregion

        #region 設定の保存と読み込み

        private void SaveSettings()
        {
            try
            {
                AppSettings settings = new AppSettings
                {
                    KeywordHistory = new List<string>(),
                    FolderPathHistory = new List<string>(),
                    UseRegex = chkUseRegex.Checked
                };

                // キーワード履歴を保存
                foreach (var item in cmbKeyword.Items)
                {
                    settings.KeywordHistory.Add(item.ToString());
                }

                // フォルダパス履歴を保存
                foreach (var item in cmbFolderPath.Items)
                {
                    settings.FolderPathHistory.Add(item.ToString());
                }

                // JavaScriptSerializerを使用してJSONに変換
                JavaScriptSerializer serializer = new JavaScriptSerializer();
                string json = serializer.Serialize(settings);
                
                // JSONファイルとして保存
                File.WriteAllText(GetSettingsFilePath(), json, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                // 保存エラーは無視（ユーザーに通知しない）
                Console.WriteLine($"設定の保存中にエラーが発生しました: {ex.Message}");
            }
        }

        private void LoadSettings()
        {
            string settingsFilePath = GetSettingsFilePath();
            
            if (!File.Exists(settingsFilePath))
            {
                return; // 設定ファイルが存在しない場合は何もしない
            }

            try
            {
                string json = File.ReadAllText(settingsFilePath, Encoding.UTF8);
                
                // JavaScriptSerializerを使用してJSONをデシリアライズ
                JavaScriptSerializer serializer = new JavaScriptSerializer();
                AppSettings settings = serializer.Deserialize<AppSettings>(json);

                if (settings != null)
                {
                    // キーワード履歴を読み込み
                    cmbKeyword.Items.Clear();
                    foreach (var item in settings.KeywordHistory)
                    {
                        cmbKeyword.Items.Add(item);
                    }

                    // フォルダパス履歴を読み込み
                    cmbFolderPath.Items.Clear();
                    foreach (var item in settings.FolderPathHistory)
                    {
                        cmbFolderPath.Items.Add(item);
                    }

                    // 正規表現の設定を読み込み
                    chkUseRegex.Checked = settings.UseRegex;

                    // 最新の項目をテキストボックスに表示
                    if (cmbKeyword.Items.Count > 0)
                    {
                        cmbKeyword.SelectedIndex = 0;
                    }

                    if (cmbFolderPath.Items.Count > 0)
                    {
                        cmbFolderPath.SelectedIndex = 0;
                    }
                }
            }
            catch (Exception ex)
            {
                // 読み込みエラーは無視（ユーザーに通知しない）
                Console.WriteLine($"設定の読み込み中にエラーが発生しました: {ex.Message}");
            }
        }

        private string GetSettingsFilePath()
        {
            return Path.Combine(Application.StartupPath, SettingsFileName);
        }

        #endregion

        #region エクスプローラーでファイルを選択（ツリー展開）

        private static void OpenFolderAndSelectFile(string filePath)
        {
            if (File.Exists(filePath))
            {
                try
                {
                    // 最もシンプルで確実な方法
                    Process.Start("explorer.exe", "/select,\"" + filePath + "\"");
                }
                catch (Exception ex)
                {
                    // エラーが発生した場合は、フォルダだけを開く
                    try
                    {
                        string folderPath = Path.GetDirectoryName(filePath);
                        if (Directory.Exists(folderPath))
                        {
                            Process.Start("explorer.exe", folderPath);
                        }
                    }
                    catch (Exception innerEx)
                    {
                        MessageBox.Show($"フォルダを開けませんでした: {innerEx.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        #endregion
    }
}
