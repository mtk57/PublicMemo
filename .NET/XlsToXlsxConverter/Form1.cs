using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace XlsToXlsxConverter
{
    public partial class MainForm : Form
    {
        private CancellationTokenSource _cancellationTokenSource;
        private string _settingsFilePath = "converter_settings.json";
        private bool _isConverting = false;
        private const int MaxHistoryItems = 10;

        // モデルクラス - 設定の保存/読み込み用
        [DataContract]
        public class Settings
        {
            [DataMember]
            public string FolderPath { get; set; } = "";
            
            [DataMember]
            public bool DeleteOriginalFiles { get; set; } = false;
            
            [DataMember]
            public List<string> FolderPathHistory { get; set; } = new List<string>();
            
            [DataMember]
            public bool OverwriteExistingFiles { get; set; } = true;
        }

        // 変換結果を格納するクラス
        public class ConversionResult
        {
            public string SourceFilePath { get; set; }
            public string DestinationFilePath { get; set; }
            public bool Success { get; set; }
            public string ErrorMessage { get; set; }
        }

        public MainForm()
        {
            InitializeComponent();
            this.FormClosing += MainForm_FormClosing;
            this.Load += MainForm_Load;
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            LoadSettings();
        }

        private void LoadSettings()
        {
            try
            {
                if (File.Exists(_settingsFilePath))
                {
                    // DataContractJsonSerializerを使用してJSON読み込み
                    using (FileStream fs = new FileStream(_settingsFilePath, FileMode.Open))
                    {
                        DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(Settings));
                        Settings settings = (Settings)serializer.ReadObject(fs);

                        // コンボボックスの履歴をクリアして設定
                        cmbFolderPath.Items.Clear();
                        foreach (var path in settings.FolderPathHistory)
                        {
                            cmbFolderPath.Items.Add(path);
                        }
                        cmbFolderPath.Text = settings.FolderPath;

                        // その他の設定
                        chkDeleteOriginal.Checked = settings.DeleteOriginalFiles;
                        chkOverwrite.Checked = settings.OverwriteExistingFiles;
                    }
                }
            }
            catch (Exception ex)
            {
                // 読み込み失敗時は何もしない
                Console.WriteLine($"設定の読み込みに失敗: {ex.Message}");
            }
        }

        private void SaveSettings()
        {
            try
            {
                // フォルダパス履歴を更新
                List<string> folderPathHistory = new List<string>();
                if (!string.IsNullOrEmpty(cmbFolderPath.Text))
                {
                    folderPathHistory.Add(cmbFolderPath.Text);
                }
                foreach (var item in cmbFolderPath.Items)
                {
                    string path = item.ToString();
                    if (!folderPathHistory.Contains(path) && folderPathHistory.Count < MaxHistoryItems)
                    {
                        folderPathHistory.Add(path);
                    }
                }

                // 設定を保存
                Settings settings = new Settings
                {
                    FolderPath = cmbFolderPath.Text,
                    DeleteOriginalFiles = chkDeleteOriginal.Checked,
                    FolderPathHistory = folderPathHistory,
                    OverwriteExistingFiles = chkOverwrite.Checked
                };

                // DataContractJsonSerializerを使用してJSON保存
                using (FileStream fs = new FileStream(_settingsFilePath, FileMode.Create))
                {
                    DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(Settings));
                    serializer.WriteObject(fs, settings);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"設定の保存に失敗しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 履歴コンボボックスに項目を追加（重複なしで先頭に配置）
        private void AddToComboBoxHistory(ComboBox comboBox, string item)
        {
            if (string.IsNullOrEmpty(item))
                return;

            // 既存の項目を削除（重複を防ぐ）
            if (comboBox.Items.Contains(item))
            {
                comboBox.Items.Remove(item);
            }

            // 先頭に追加
            comboBox.Items.Insert(0, item);

            // 最大履歴数を超えた場合、古い項目を削除
            while (comboBox.Items.Count > MaxHistoryItems)
            {
                comboBox.Items.RemoveAt(comboBox.Items.Count - 1);
            }

            // 現在の選択項目を設定
            comboBox.Text = item;
        }

        private void btnSelectFolder_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                dialog.Description = "変換するXLSファイルがあるフォルダを選択してください";
                dialog.ShowNewFolderButton = false;

                if (!string.IsNullOrEmpty(cmbFolderPath.Text) && Directory.Exists(cmbFolderPath.Text))
                {
                    dialog.SelectedPath = cmbFolderPath.Text;
                }

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    AddToComboBoxHistory(cmbFolderPath, dialog.SelectedPath);
                }
            }
        }

        private async void btnStartConversion_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(cmbFolderPath.Text))
            {
                MessageBox.Show("フォルダパスを入力してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!Directory.Exists(cmbFolderPath.Text))
            {
                MessageBox.Show("指定されたフォルダが存在しません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // UI状態の設定
            SetConvertingState(true);

            // キャンセルトークンを作成
            _cancellationTokenSource = new CancellationTokenSource();

            try
            {
                // 結果リストをクリア
                lstResults.Items.Clear();

                // 変換処理を実行
                await ConvertXlsFilesAsync(cmbFolderPath.Text, _cancellationTokenSource.Token);

                lblStatus.Text = "変換完了";
            }
            catch (OperationCanceledException)
            {
                lblStatus.Text = "変換は中止されました";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"変換中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                lblStatus.Text = "エラーが発生しました";
            }
            finally
            {
                // UI状態を通常に戻す
                SetConvertingState(false);
                progressBar.Value = 0;
            }
        }

        private void btnCancelConversion_Click(object sender, EventArgs e)
        {
            if (_isConverting && _cancellationTokenSource != null)
            {
                _cancellationTokenSource.Cancel();
                lblStatus.Text = "キャンセル処理中...";
            }
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 変換中の場合、キャンセルする
            if (_isConverting && _cancellationTokenSource != null)
            {
                _cancellationTokenSource.Cancel();
            }

            // 設定を保存
            SaveSettings();
        }

        private void SetConvertingState(bool isConverting)
        {
            _isConverting = isConverting;

            // 変換中はUIの一部を無効化
            cmbFolderPath.Enabled = !isConverting;
            chkDeleteOriginal.Enabled = !isConverting;
            chkOverwrite.Enabled = !isConverting;
            btnSelectFolder.Enabled = !isConverting;
            btnStartConversion.Enabled = !isConverting;
            btnCancelConversion.Enabled = isConverting;

            // 変換中はステータスを更新
            if (isConverting)
            {
                lblStatus.Text = "変換中...";
            }
        }

        private async Task ConvertXlsFilesAsync(string folderPath, CancellationToken cancellationToken)
        {
            // 対象のXLSファイルを取得
            string[] xlsFiles = Directory.GetFiles(folderPath, "*.xls", SearchOption.AllDirectories)
                .Where(f => !f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) && 
                            !f.Contains("~$") && // 一時ファイルを除外
                            Path.GetExtension(f).Equals(".xls", StringComparison.OrdinalIgnoreCase))
                .ToArray();

            // 進捗状況の初期化
            int totalFiles = xlsFiles.Length;
            if (totalFiles == 0)
            {
                UpdateStatus("変換対象のXLSファイルが見つかりませんでした");
                return;
            }

            UpdateProgressBar(0, totalFiles);
            int processedFiles = 0;
            int successCount = 0;
            int failCount = 0;

            // Excel Applicationの初期化
            Excel.Application excelApp = null;
            try
            {
                excelApp = new Excel.Application();
                excelApp.DisplayAlerts = false; // ダイアログを表示しない

                // 各ファイルを順に処理
                foreach (string xlsFilePath in xlsFiles)
                {
                    // キャンセル確認
                    if (cancellationToken.IsCancellationRequested)
                    {
                        throw new OperationCanceledException();
                    }

                    // XLSXファイル名の生成
                    string xlsxFilePath = Path.ChangeExtension(xlsFilePath, ".xlsx");

                    try
                    {
                        // 変換処理開始
                        UpdateStatus($"変換中: {Path.GetFileName(xlsFilePath)}");
                        
                        // 既存ファイルの確認
                        if (File.Exists(xlsxFilePath) && !chkOverwrite.Checked)
                        {
                            AddToResultsList(new ConversionResult
                            {
                                SourceFilePath = xlsFilePath,
                                DestinationFilePath = xlsxFilePath,
                                Success = false,
                                ErrorMessage = "ファイルが既に存在します（上書きオプションがオフ）"
                            });
                            failCount++;
                            continue;
                        }

                        // ファイルを変換
                        await Task.Run(() => 
                        {
                            ConvertXlsToXlsx(excelApp, xlsFilePath, xlsxFilePath, cancellationToken);
                        }, cancellationToken);

                        // 元ファイルの削除（オプション）
                        if (chkDeleteOriginal.Checked)
                        {
                            File.Delete(xlsFilePath);
                        }

                        // 結果の追加
                        AddToResultsList(new ConversionResult
                        {
                            SourceFilePath = xlsFilePath,
                            DestinationFilePath = xlsxFilePath,
                            Success = true
                        });
                        successCount++;
                    }
                    catch (OperationCanceledException)
                    {
                        throw;
                    }
                    catch (Exception ex)
                    {
                        // エラー情報を結果に追加
                        AddToResultsList(new ConversionResult
                        {
                            SourceFilePath = xlsFilePath,
                            DestinationFilePath = xlsxFilePath,
                            Success = false,
                            ErrorMessage = ex.Message
                        });
                        failCount++;
                    }

                    // 進捗の更新
                    processedFiles++;
                    UpdateProgressBar(processedFiles, totalFiles);
                    UpdateStatus($"処理中... {processedFiles}/{totalFiles} ファイル");
                }

                // 完了メッセージ
                UpdateStatus($"変換完了: 成功={successCount}, 失敗={failCount}, 合計={totalFiles}");
            }
            finally
            {
                // Excel Applicationの終了
                if (excelApp != null)
                {
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                }
            }
        }

        private void ConvertXlsToXlsx(Excel.Application excelApp, string sourcePath, string destinationPath, CancellationToken cancellationToken)
        {
            Excel.Workbook workbook = null;
            try
            {
                // ファイルを開く
                workbook = excelApp.Workbooks.Open(
                    sourcePath,
                    UpdateLinks: 0,
                    ReadOnly: true,
                    Format: 5,  // XLS形式
                    Password: Type.Missing,
                    WriteResPassword: Type.Missing,
                    IgnoreReadOnlyRecommended: true,
                    Origin: Type.Missing,
                    Delimiter: Type.Missing,
                    Editable: false,
                    Notify: false,
                    Converter: Type.Missing,
                    AddToMru: false,
                    Local: true,
                    CorruptLoad: false
                );

                if (cancellationToken.IsCancellationRequested)
                {
                    throw new OperationCanceledException();
                }

                // XLSX形式で保存
                workbook.SaveAs(
                    destinationPath,
                    Excel.XlFileFormat.xlOpenXMLWorkbook,  // XLSX形式
                    Type.Missing,
                    Type.Missing,
                    false,  // ReadOnlyRecommended
                    false,  // CreateBackup
                    Excel.XlSaveAsAccessMode.xlNoChange,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing
                );
            }
            finally
            {
                // リソースの解放
                if (workbook != null)
                {
                    workbook.Close(false);
                    Marshal.ReleaseComObject(workbook);
                }
            }
        }

        private void UpdateProgressBar(int current, int total)
        {
            if (progressBar.InvokeRequired)
            {
                progressBar.Invoke(new Action(() => {
                    progressBar.Maximum = total;
                    progressBar.Value = current > total ? total : current;
                }));
            }
            else
            {
                progressBar.Maximum = total;
                progressBar.Value = current > total ? total : current;
            }
        }

        private void UpdateStatus(string message)
        {
            if (lblStatus.InvokeRequired)
            {
                lblStatus.Invoke(new Action(() => lblStatus.Text = message));
            }
            else
            {
                lblStatus.Text = message;
            }
        }

        private void AddToResultsList(ConversionResult result)
        {
            if (lstResults.InvokeRequired)
            {
                lstResults.Invoke(new Action(() => AddResultToList(result)));
            }
            else
            {
                AddResultToList(result);
            }
        }

        private void AddResultToList(ConversionResult result)
        {
            string status = result.Success ? "成功" : "失敗";
            string errorInfo = result.Success ? "" : $" - エラー: {result.ErrorMessage}";
            
            ListViewItem item = new ListViewItem(new string[] {
                Path.GetFileName(result.SourceFilePath),
                status,
                result.SourceFilePath,
                result.DestinationFilePath,
                errorInfo
            });
            
            item.ForeColor = result.Success ? System.Drawing.Color.Green : System.Drawing.Color.Red;
            lstResults.Items.Add(item);
        }

        private void lstResults_DoubleClick(object sender, EventArgs e)
        {
            if (lstResults.SelectedItems.Count > 0)
            {
                ListViewItem item = lstResults.SelectedItems[0];
                string folderPath = Path.GetDirectoryName(item.SubItems[2].Text);
                
                if (!string.IsNullOrEmpty(folderPath) && Directory.Exists(folderPath))
                {
                    System.Diagnostics.Process.Start("explorer.exe", folderPath);
                }
            }
        }
    }
}