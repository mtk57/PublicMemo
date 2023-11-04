using System;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace ClipboardImageSaver
{
    /// <summary>
    /// クリップボードに画像が入ったら、指定フォルダに指定ファイル名(連番)で保存する(PNG形式)
    /// </summary>
    public partial class MainForm : Form
    {
        [DllImport("user32.dll", SetLastError = true)]
        private extern static void AddClipboardFormatListener(IntPtr hwnd);

        [DllImport("user32.dll", SetLastError = true)]
        private extern static void RemoveClipboardFormatListener(IntPtr hwnd);

        private const int WM_CLIPBOARDUPDATE = 0x31D;
        private const string CaptionResume = "クリップボード監視を一時停止する";
        private const string CaptionSuspend = "クリップボード監視を再開する";
        private const string Ext = ".png";

        private string _settingsJsonPath = string.Empty;
        private Settings _settings = null;
        private bool _isResume = true;
        private decimal _num = 0;

        private string SavePath
        {
            get
            {
                return $"{_settings.SaveDirPath}\\{_settings.SaveFileName}{_num}{Ext}";
            }
        }

        public MainForm()
        {
            InitializeComponent();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            try
            {
                LoadSettings();

                WriteLog(_settings.ToString());
            }
            catch(Exception ex)
            {
                ShowException(ex);
                _settings = GetEmptySettings();
                CopyFromSettingsToUI();
            }
        }

        private void MainForm_Shown(object sender, EventArgs e)
        {
            if (_settings == null) return;

            AddClipboardFormatListener(Handle);
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (_settings == null) return;

            RemoveClipboardFormatListener(Handle);

            try
            {
                SaveSettings();
            }
            catch (Exception ex)
            {
                ShowException(ex);
            }
        }

        protected override void WndProc(ref Message m)
        {
            if (!_isResume &&
                m.Msg == WM_CLIPBOARDUPDATE &&
                !string.IsNullOrEmpty(_settings.SaveDirPath) &&
                Directory.Exists(_settings.SaveDirPath) &&
                !string.IsNullOrEmpty(_settings.SaveFileName) &&
                Clipboard.ContainsImage()
                )
            {
                var img = Clipboard.GetImage();
                if (img == null)
                {
                    base.WndProc(ref m);
                    return;
                }

                CopyFromUiToSettings();

                _num = _settings.StartNum;
                string savePath;

                while (true)
                {
                    savePath = SavePath;

                    if (File.Exists(savePath))
                    {
                        _num++;
                    }
                    else
                    {
                        break;
                    }
                }

                try
                {
                    img.Save(savePath, ImageFormat.Png);
                    Clipboard.Clear();
                    m.Result = IntPtr.Zero;

                    WriteLog(savePath);
                }
                catch(Exception ex)
                {
                    WriteLog($"Error!! Msg={ex.Message}, Stack={ex.StackTrace}, SavePath={savePath}");
                }
            }
            else
            {
                base.WndProc(ref m);
            }

        }//WndProc

        private void ButtonSaveDirRef_Click(object sender, EventArgs e)
        {
            var fbd = new FolderBrowserDialog
            {
                Description = "保存フォルダを指定してください。",
                RootFolder = Environment.SpecialFolder.Desktop,
                SelectedPath = @"C:\Windows",
                ShowNewFolderButton = true
            };

            if (fbd.ShowDialog(this) == DialogResult.OK)
            {
                _settings.SaveDirPath = fbd.SelectedPath;
                textBoxSaveDirPath.Text = fbd.SelectedPath;
            }
        }

        private void ButtonResume_Click(object sender, EventArgs e)
        {
            if (_isResume)
            {
                buttonResume.Text = CaptionResume;
                WriteLog("-------------------------------");
                WriteLog("●START");
            }
            else
            {
                buttonResume.Text = CaptionSuspend;
                WriteLog("●PAUSE");
            }

            CopyFromUiToSettings();
            WriteLog(_settings.ToString());

            _isResume = !_isResume;
        }

        private void LoadSettings()
        {
            var myDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            _settingsJsonPath = Path.Combine(myDir, "Settings.json");
            if (!File.Exists(_settingsJsonPath))
            {
                _settings = GetEmptySettings();
            }
            else
            {
                _settings = ReadSettingsJson();
            }
            CopyFromSettingsToUI();
        }

        private void SaveSettings()
        {
            CopyFromUiToSettings();
            WriteSettingsJson(_settings);
        }

        private Settings ReadSettingsJson()
        {
            var json = File.ReadAllText(_settingsJsonPath);
            return SettingUtil.Deserialize<Settings>(json);
        }

        private void WriteSettingsJson(Settings settings)
        {
            var serialize = SettingUtil.Serialize<Settings>(settings);
            File.WriteAllText(_settingsJsonPath, serialize);
        }

        private Settings GetEmptySettings()
        {
            return new Settings { SaveDirPath = "", SaveFileName = "", StartNum = 0 };
        }

        private void CopyFromSettingsToUI()
        {
            textBoxSaveDirPath.Text = _settings.SaveDirPath;
            textBoxSaveFileName.Text = _settings.SaveFileName;
            numericUpDownStartNum.Value = _settings.StartNum;
        }

        private void CopyFromUiToSettings()
        {
            _settings.SaveDirPath = textBoxSaveDirPath.Text;
            _settings.SaveFileName = textBoxSaveFileName.Text;
            _settings.StartNum = numericUpDownStartNum.Value;
        }

        private void ShowException(Exception ex)
        {
            var msg = $"Msg={ex.Message}, Stack={ex.StackTrace}";
            MessageBox.Show(msg);
            WriteLog(msg);
        }

        private void WriteLog(string msg)
        {
            textBoxLog.AppendText(msg + Environment.NewLine);

            textBoxLog.SelectionStart = textBoxLog.Text.Length;
            textBoxLog.Focus();
            textBoxLog.ScrollToCaret();
        }
    }
}
