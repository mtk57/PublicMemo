using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json; // 標準のJSONシリアライザ
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SimpleExcelGrep
{
    public partial class MainForm : Form
    {
        private CancellationTokenSource _cancellationTokenSource;
        private string _settingsFilePath = "settings.json";
        private bool _isSearching = false;
        private const int MaxHistoryItems = 10;
        private List<SearchResult> _searchResults = new List<SearchResult>(); // 検索結果を保持するリスト

        // モデルクラス - 設定の保存/読み込み用
        [DataContract]
        public class Settings
        {
            [DataMember]
            public string FolderPath { get; set; } = "";
            
            [DataMember]
            public string SearchKeyword { get; set; } = "";
            
            [DataMember]
            public bool UseRegex { get; set; } = false;
            
            [DataMember]
            public string IgnoreKeywords { get; set; } = "";

            [DataMember]
            public List<string> FolderPathHistory { get; set; } = new List<string>();

            [DataMember]
            public List<string> SearchKeywordHistory { get; set; } = new List<string>();
            
            [DataMember]
            public List<string> IgnoreKeywordsHistory { get; set; } = new List<string>();
            
            [DataMember]
            public bool RealTimeDisplay { get; set; } = true; // 追加: リアルタイム表示設定
        }

        // 検索結果を格納するクラス
        public class SearchResult
        {
            public string FilePath { get; set; }
            public string SheetName { get; set; }
            public string CellPosition { get; set; }
            public string CellValue { get; set; }
        }

        public MainForm()
        {
            InitializeComponent();
            this.FormClosing += MainForm_FormClosing;
            this.Load += MainForm_Load;
            btnSelectFolder.Click += BtnSelectFolder_Click;
            btnStartSearch.Click += BtnStartSearch_Click;
            btnCancelSearch.Click += BtnCancelSearch_Click;
            grdResults.DoubleClick += GrdResults_DoubleClick;
            
            // グリッドにキー押下イベントを追加
            grdResults.KeyDown += GrdResults_KeyDown;
            
            // 複数行選択を可能にする
            grdResults.MultiSelect = true;
            grdResults.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            
            // コンテキストメニューを追加
            ContextMenuStrip contextMenu = new ContextMenuStrip();
            ToolStripMenuItem copyMenuItem = new ToolStripMenuItem("コピー");
            copyMenuItem.Click += (s, e) => CopySelectedRowsToClipboard();
            contextMenu.Items.Add(copyMenuItem);
            
            ToolStripMenuItem selectAllMenuItem = new ToolStripMenuItem("すべて選択");
            selectAllMenuItem.Click += (s, e) => grdResults.SelectAll();
            contextMenu.Items.Add(selectAllMenuItem);
            
            grdResults.ContextMenuStrip = contextMenu;
            
            // リアルタイム表示チェックボックスの状態変更イベントを登録
            chkRealTimeDisplay.CheckedChanged += (s, e) => {
                // 設定を保存
                SaveSettings();
            };
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

                        // 検索キーワード履歴
                        cmbKeyword.Items.Clear();
                        foreach (var keyword in settings.SearchKeywordHistory)
                        {
                            cmbKeyword.Items.Add(keyword);
                        }
                        cmbKeyword.Text = settings.SearchKeyword;

                        // 無視キーワード履歴
                        cmbIgnoreKeywords.Items.Clear();
                        foreach (var keyword in settings.IgnoreKeywordsHistory)
                        {
                            cmbIgnoreKeywords.Items.Add(keyword);
                        }
                        cmbIgnoreKeywords.Text = settings.IgnoreKeywords;

                        // その他の設定
                        chkRegex.Checked = settings.UseRegex;
                        
                        // リアルタイム表示設定
                        chkRealTimeDisplay.Checked = settings.RealTimeDisplay;
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

                // 検索キーワード履歴を更新
                List<string> searchKeywordHistory = new List<string>();
                if (!string.IsNullOrEmpty(cmbKeyword.Text))
                {
                    searchKeywordHistory.Add(cmbKeyword.Text);
                }
                foreach (var item in cmbKeyword.Items)
                {
                    string keyword = item.ToString();
                    if (!searchKeywordHistory.Contains(keyword) && searchKeywordHistory.Count < MaxHistoryItems)
                    {
                        searchKeywordHistory.Add(keyword);
                    }
                }
                
                // 無視キーワード履歴を更新
                List<string> ignoreKeywordsHistory = new List<string>();
                if (!string.IsNullOrEmpty(cmbIgnoreKeywords.Text))
                {
                    ignoreKeywordsHistory.Add(cmbIgnoreKeywords.Text);
                }
                foreach (var item in cmbIgnoreKeywords.Items)
                {
                    string keyword = item.ToString();
                    if (!ignoreKeywordsHistory.Contains(keyword) && ignoreKeywordsHistory.Count < MaxHistoryItems)
                    {
                        ignoreKeywordsHistory.Add(keyword);
                    }
                }

                // 設定を保存
                Settings settings = new Settings
                {
                    FolderPath = cmbFolderPath.Text,
                    SearchKeyword = cmbKeyword.Text,
                    UseRegex = chkRegex.Checked,
                    IgnoreKeywords = cmbIgnoreKeywords.Text,
                    FolderPathHistory = folderPathHistory,
                    SearchKeywordHistory = searchKeywordHistory,
                    IgnoreKeywordsHistory = ignoreKeywordsHistory,
                    RealTimeDisplay = chkRealTimeDisplay.Checked // リアルタイム表示設定を保存
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

        private void BtnSelectFolder_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                dialog.Description = "検索するフォルダを選択してください";
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

        private async void BtnStartSearch_Click(object sender, EventArgs e)
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

            if (string.IsNullOrWhiteSpace(cmbKeyword.Text))
            {
                MessageBox.Show("検索キーワードを入力してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // 検索キーワードを履歴に追加
            AddToComboBoxHistory(cmbKeyword, cmbKeyword.Text);
            
            // 無視キーワードを履歴に追加
            AddToComboBoxHistory(cmbIgnoreKeywords, cmbIgnoreKeywords.Text);

            // UIを検索中の状態に変更
            SetSearchingState(true);

            // キャンセルトークンを作成
            _cancellationTokenSource = new CancellationTokenSource();
            
            // 検索結果リストをクリア
            _searchResults.Clear();

            try
            {
                // 結果グリッドをクリア
                grdResults.Rows.Clear();

                // 無視キーワードのリストを作成
                List<string> ignoreKeywords = cmbIgnoreKeywords.Text
                    .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(k => k.Trim())
                    .Where(k => !string.IsNullOrEmpty(k))
                    .ToList();

                // 正規表現オブジェクト
                Regex regex = null;
                if (chkRegex.Checked)
                {
                    try
                    {
                        regex = new Regex(cmbKeyword.Text, RegexOptions.Compiled);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"正規表現が無効です: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        SetSearchingState(false);
                        return;
                    }
                }

                // 検索処理を実行
                bool isRealTimeDisplay = chkRealTimeDisplay.Checked;
                
                // 検索結果を取得
                List<SearchResult> results = await SearchExcelFilesAsync(
                    cmbFolderPath.Text,
                    cmbKeyword.Text,
                    chkRegex.Checked,
                    regex,
                    ignoreKeywords,
                    isRealTimeDisplay,
                    _cancellationTokenSource.Token);

                // リアルタイム表示がOFFの場合または検索が途中でキャンセルされた場合に、
                // 最終的な結果をまとめて表示
                if (!isRealTimeDisplay || _cancellationTokenSource.IsCancellationRequested)
                {
                    DisplaySearchResults(_searchResults);
                }

                lblStatus.Text = $"検索完了: {_searchResults.Count} 件見つかりました";
            }
            catch (OperationCanceledException)
            {
                lblStatus.Text = $"検索は中止されました: {_searchResults.Count} 件見つかりました";
                
                // キャンセル時は現在までの結果を表示
                DisplaySearchResults(_searchResults);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"検索中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                lblStatus.Text = "エラーが発生しました";
                
                // エラー時は現在までの結果を表示
                DisplaySearchResults(_searchResults);
            }
            finally
            {
                // UIを通常の状態に戻す
                SetSearchingState(false);
            }
        }

        // 検索結果をグリッドに表示するメソッド
        private void DisplaySearchResults(List<SearchResult> results)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() => DisplaySearchResults(results)));
                return;
            }
            
            // 結果グリッドをクリア
            grdResults.Rows.Clear();
            
            // 結果をグリッドに表示
            foreach (var result in results)
            {
                string fileName = Path.GetFileName(result.FilePath);
                grdResults.Rows.Add(result.FilePath, fileName, result.SheetName, result.CellPosition, result.CellValue);
            }
        }
        
        // 検索結果を1件追加するメソッド（リアルタイム表示用）
        private void AddSearchResult(SearchResult result)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() => AddSearchResult(result)));
                return;
            }
            
            string fileName = Path.GetFileName(result.FilePath);
            grdResults.Rows.Add(result.FilePath, fileName, result.SheetName, result.CellPosition, result.CellValue);
            
            // 最新の行にスクロール
            if (grdResults.Rows.Count > 0)
            {
                grdResults.FirstDisplayedScrollingRowIndex = grdResults.Rows.Count - 1;
            }
        }

        private void BtnCancelSearch_Click(object sender, EventArgs e)
        {
            if (_isSearching && _cancellationTokenSource != null)
            {
                _cancellationTokenSource.Cancel();
                lblStatus.Text = "キャンセル処理中...";
            }
        }

        private void GrdResults_DoubleClick(object sender, EventArgs e)
        {
            if (grdResults.SelectedRows.Count > 0)
            {
                string filePath = grdResults.SelectedRows[0].Cells[0].Value.ToString();
                if (!string.IsNullOrEmpty(filePath) && File.Exists(filePath))
                {
                    string folderPath = Path.GetDirectoryName(filePath);
                    System.Diagnostics.Process.Start("explorer.exe", folderPath);
                }
            }
        }
        
        // GridViewのキーダウンイベントハンドラ
        private void GrdResults_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.A)
            {
                // CTRL+A で全行選択
                grdResults.SelectAll();
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
            else if (e.Control && e.KeyCode == Keys.C)
            {
                // CTRL+C でコピー
                CopySelectedRowsToClipboard();
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }
        
        // 選択行をクリップボードにコピーするメソッド
        private void CopySelectedRowsToClipboard()
        {
            if (grdResults.SelectedRows.Count > 0)
            {
                StringBuilder sb = new StringBuilder();
                
                // ヘッダー行を追加
                for (int i = 0; i < grdResults.Columns.Count; i++)
                {
                    sb.Append(grdResults.Columns[i].HeaderText);
                    sb.Append(i == grdResults.Columns.Count - 1 ? Environment.NewLine : "\t");
                }
                
                // 選択行を追加（選択順に処理）
                List<DataGridViewRow> selectedRows = new List<DataGridViewRow>();
                foreach (DataGridViewRow row in grdResults.SelectedRows)
                {
                    selectedRows.Add(row);
                }
                
                // インデックスでソート（上から下の順番になるように）
                selectedRows.Sort((x, y) => x.Index.CompareTo(y.Index));
                
                foreach (DataGridViewRow row in selectedRows)
                {
                    for (int i = 0; i < grdResults.Columns.Count; i++)
                    {
                        sb.Append(row.Cells[i].Value?.ToString() ?? "");
                        sb.Append(i == grdResults.Columns.Count - 1 ? Environment.NewLine : "\t");
                    }
                }
                
                try
                {
                    Clipboard.SetText(sb.ToString());
                    // コピー成功を通知（オプション）
                    lblStatus.Text = $"{selectedRows.Count}行をクリップボードにコピーしました";
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"クリップボードへのコピーに失敗しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 検索中の場合、キャンセルする
            if (_isSearching && _cancellationTokenSource != null)
            {
                _cancellationTokenSource.Cancel();
            }

            // 設定を保存
            SaveSettings();
        }

        private void SetSearchingState(bool isSearching)
        {
            _isSearching = isSearching;

            // 検索中はUIの一部を無効化
            cmbFolderPath.Enabled = !isSearching;
            cmbKeyword.Enabled = !isSearching;
            cmbIgnoreKeywords.Enabled = !isSearching;
            chkRegex.Enabled = !isSearching;
            btnSelectFolder.Enabled = !isSearching;
            btnStartSearch.Enabled = !isSearching;
            btnCancelSearch.Enabled = isSearching;
            
            // リアルタイム表示チェックボックスは検索中も有効のまま
            // chkRealTimeDisplay.Enabled は変更しない

            // 検索中はステータスを更新
            if (isSearching)
            {
                lblStatus.Text = "検索中...";
            }
        }

        private async Task<List<SearchResult>> SearchExcelFilesAsync(
            string folderPath,
            string keyword,
            bool useRegex,
            Regex regex,
            List<string> ignoreKeywords,
            bool isRealTimeDisplay,
            CancellationToken cancellationToken)
        {
            // 結果グリッドをクリア
            if (isRealTimeDisplay)
            {
                this.Invoke(new Action(() => grdResults.Rows.Clear()));
            }

            // Excelファイルの一覧を取得
            string[] excelFiles = Directory.GetFiles(folderPath, "*.xls*", SearchOption.AllDirectories)
                .Where(f => f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) || 
                            f.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
                .ToArray();

            // 進捗状況を表示するための変数
            int totalFiles = excelFiles.Length;
            int processedFiles = 0;

            // 各ファイルを処理
            foreach (string filePath in excelFiles)
            {
                // キャンセルされた場合は処理を終了
                if (cancellationToken.IsCancellationRequested)
                {
                    throw new OperationCanceledException();
                }

                // 無視キーワードが含まれている場合はスキップ
                if (ignoreKeywords.Any(k => filePath.Contains(k)))
                {
                    processedFiles++;
                    UpdateStatus($"処理中... {processedFiles}/{totalFiles} ファイル");
                    continue;
                }

                try
                {
                    // ファイル拡張子によって処理を分ける
                    string extension = Path.GetExtension(filePath).ToLower();
                    List<SearchResult> fileResults = new List<SearchResult>();

                    if (extension == ".xlsx")
                    {
                        // .xlsx ファイルはOpenXMLで処理
                        fileResults = await Task.Run(() => SearchInXlsxFile(
                            filePath, keyword, useRegex, regex, isRealTimeDisplay, cancellationToken), cancellationToken);
                    }
                    else if (extension == ".xls")
                    {
                        // .xls ファイルはサポート外として処理（メッセージのみ表示）
                        UpdateStatus($"注: .xls形式は現在サポートされていません: {filePath}");
                        processedFiles++;
                        continue;
                    }

                    // 結果を追加
                    _searchResults.AddRange(fileResults);

                    // 進捗状況を更新
                    processedFiles++;
                    UpdateStatus($"処理中... {processedFiles}/{totalFiles} ファイル ({_searchResults.Count} 件見つかりました)");
                }
                catch (OperationCanceledException)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    // ファイル処理中のエラーをログに記録するなどの処理
                    Console.WriteLine($"ファイル処理エラー: {filePath}, {ex.Message}");
                }
            }

            return _searchResults;
        }

        private List<SearchResult> SearchInXlsxFile(
            string filePath,
            string keyword,
            bool useRegex,
            Regex regex,
            bool isRealTimeDisplay,
            CancellationToken cancellationToken)
        {
            List<SearchResult> results = new List<SearchResult>();

            try
            {
                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
                {
                    WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                    SharedStringTablePart sharedStringTablePart = workbookPart.SharedStringTablePart;
                    SharedStringTable sharedStringTable = sharedStringTablePart?.SharedStringTable;

                    // ワークシートの一覧を取得
                    Sheets sheets = workbookPart.Workbook.Sheets;
                    
                    // 各シートを処理
                    foreach (Sheet sheet in sheets.Elements<Sheet>())
                    {
                        // キャンセル処理
                        if (cancellationToken.IsCancellationRequested)
                        {
                            throw new OperationCanceledException();
                        }

                        // シートIDを取得
                        string relationshipId = sheet.Id.Value;
                        WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(relationshipId);
                        SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                        // 各行を処理
                        foreach (Row row in sheetData.Elements<Row>())
                        {
                            // キャンセル処理
                            if (cancellationToken.IsCancellationRequested)
                            {
                                throw new OperationCanceledException();
                            }

                            // 各セルを処理
                            foreach (Cell cell in row.Elements<Cell>())
                            {
                                // キャンセル処理
                                if (cancellationToken.IsCancellationRequested)
                                {
                                    throw new OperationCanceledException();
                                }

                                // セルの値を取得
                                string cellValue = GetCellValue(cell, sharedStringTable);
                                
                                if (!string.IsNullOrEmpty(cellValue))
                                {
                                    bool isMatch;
                                    
                                    if (useRegex && regex != null)
                                    {
                                        isMatch = regex.IsMatch(cellValue);
                                    }
                                    else
                                    {
                                        isMatch = cellValue.Contains(keyword);
                                    }

                                    if (isMatch)
                                    {
                                        SearchResult result = new SearchResult
                                        {
                                            FilePath = filePath,
                                            SheetName = sheet.Name,
                                            CellPosition = GetCellReference(cell),
                                            CellValue = cellValue
                                        };
                                        
                                        results.Add(result);
                                        
                                        // リアルタイム表示が有効な場合、UIに即時反映
                                        if (isRealTimeDisplay)
                                        {
                                            AddSearchResult(result);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                // エラーをログに記録
                Console.WriteLine($"Excel処理エラー: {filePath}, {ex.Message}");
            }

            return results;
        }

        // セルの値を取得するヘルパーメソッド
        private string GetCellValue(Cell cell, SharedStringTable sharedStringTable)
        {
            if (cell == null)
                return string.Empty;

            string cellValue = cell.InnerText;

            // セルの値が数値型の場合は、そのまま返す
            if (cell.DataType == null)
                return cellValue;

            // セルの値が共有文字列の場合は、共有文字列テーブルから実際の値を取得
            if (cell.DataType.Value == CellValues.SharedString && sharedStringTable != null)
            {
                int ssid = int.Parse(cellValue);
                
                if (ssid >= 0 && ssid < sharedStringTable.Count())
                {
                    SharedStringItem ssi = sharedStringTable.Elements<SharedStringItem>().ElementAt(ssid);
                    if (ssi.Text != null)
                        return ssi.Text.Text;
                    else if (ssi.InnerText != null)
                        return ssi.InnerText;
                    else if (ssi.InnerXml != null)
                        return ssi.InnerXml;
                }
            }
            
            // 日付や他の型の場合も、基本的にはInnerTextで取得できる
            return cellValue;
        }

        // セル参照（例：A1）を取得するメソッド
        private string GetCellReference(Cell cell)
        {
            return cell.CellReference?.Value ?? string.Empty;
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
    }
}