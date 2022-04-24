using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Net;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Windows.Forms;
using MyData;

/// <summary>
/// 簡易RESTサーバ
/// 
/// [memo]
/// ・exeを実行すると「アクセスが拒否されました」と出る場合は、マニフェストファイルのrequestedExecutionLevelのレベルを「requireAdministrator」に変更する。
///   参考：https://sweep3092.hatenablog.com/entry/2014/12/27/160005
/// 
/// [TODO]
/// MUST
/// ・Stopに対応しよう!
/// ・httpsに対応しよう!
/// ・Vistaで動くか試そう!
/// ・ログに対応しよう！
/// 
/// </summary>
namespace TinyRestServer
{
    public partial class Form1 : Form
    {
        private HttpListener _listener = null;
        private Dictionary<int, String> _userDatas = null;

        public Form1()
        {
            InitializeComponent();

            backgroundWorker1.WorkerReportsProgress = true;
            backgroundWorker1.WorkerSupportsCancellation = true;

            _userDatas = new Dictionary<int, string>();

            button_Stop.Enabled = false;
        }

        private void button_Start_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox_Url.Text))
            {
                MessageBox.Show("URL is nothig!");
                return;
            }
            if (string.IsNullOrEmpty(textBox_Port.Text))
            {
                MessageBox.Show("Port is nothig!");
                return;
            }

            if (backgroundWorker1.IsBusy != true)
            {
                UpdateTextBoxLog("START");

                button_Start.Enabled = false;
                button_Stop.Enabled = true;

                backgroundWorker1.RunWorkerAsync();
            }
        }

        private void button_Stop_Click(object sender, EventArgs e)
        {
            UpdateTextBoxLog("STOP");

            if (backgroundWorker1.WorkerSupportsCancellation == true)
            {
                backgroundWorker1.CancelAsync();
            }

            closeListener();

            button_Start.Enabled = true;
            button_Stop.Enabled = false;
        }

        private void buttonDefaultURL_Click(object sender, EventArgs e)
        {
            if(radioButtonHttp.Checked == true)
            {
                textBox_Url.Text = Const.DEFAULT_URL_HTTP;
                textBox_Port.Text = Const.DEFAULT_PORT_HTTP.ToString();
            }
            else
            {
                textBox_Url.Text = Const.DEFAULT_URL_HTTPS;
                textBox_Port.Text = Const.DEFAULT_PORT_HTTPS.ToString();
            }
        }

        delegate void UpdateTextBoxLogDelegate(string data);

        public void UpdateTextBoxLog(string data)
        {
            if (InvokeRequired)
            {
                Invoke(new UpdateTextBoxLogDelegate(TextBoxLogUpdate), data);
                return;
            }
            TextBoxLogUpdate(data);
        }

        private void TextBoxLogUpdate(string data)
        {
            textBox_Log.AppendText(data + Environment.NewLine);
        }

        private void startListener(string url)
        {
            closeListener();

            _listener = new HttpListener();

            UpdateTextBoxLog(url);

            _listener.Prefixes.Add(url);

            _listener.Start();
        }

        private void closeListener()
        {
            if (_listener != null)
            {
                _listener.Stop();
                _listener.Close();
                _listener = null;
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            var worker = sender as BackgroundWorker;

            startListener(textBox_Url.Text + ":" + textBox_Port.Text + @"/");

            run(e, worker);
        }

        private void run(DoWorkEventArgs e, BackgroundWorker worker)
        {
            while (_listener.IsListening)
            {
                UpdateTextBoxLog("Listening...");

                // リクエストが来るまでここで止まる
                var c = _listener.GetContext();

                if (worker.CancellationPending == true)
                {
                    e.Cancel = true;
                    break;
                }

                // リクエストを取得
                var req = c.Request;
                var res = c.Response;

                try
                {
                    UpdateTextBoxLog(req.RawUrl);
                    UpdateTextBoxLog(req.HttpMethod);

                    var reqSplit = req.RawUrl.Split('/');
                    if (reqSplit.Length < 2)
                    {
                        // URL不正
                        res.StatusCode = 404;
                        res.StatusDescription = "URL is a strange.";

                        UpdateTextBoxLog(res.StatusCode.ToString());
                        UpdateTextBoxLog(res.StatusDescription);
                        continue;
                    }

                    var api = reqSplit[1];
                    if (api != Const.API_USERS)
                    {
                        // 存在しないAPI
                        res.StatusCode = 404;
                        res.StatusDescription = "API is not exist.";

                        UpdateTextBoxLog(res.StatusCode.ToString());
                        UpdateTextBoxLog(res.StatusDescription);
                        continue;
                    }

                    if (req.HttpMethod == Const.METHOD_POST || req.HttpMethod == Const.METHOD_PUT)
                    {
                        if (req.HasEntityBody)
                        {
                            // リクエストボディに対する処理
                            using (var sr = new StreamReader(req.InputStream, new UTF8Encoding(false)))
                            {
                                var reqJson = sr.ReadToEnd();
                                UpdateTextBoxLog("[REQ]=" + reqJson);

                                var userData = Utils.Deserialize<UserData>(reqJson);

                                if (req.HttpMethod == Const.METHOD_POST)
                                {
                                    addUser(userData);
                                }
                                else
                                {
                                    updateUser(userData);
                                }
                            }
                        }

                        res.StatusCode = 200;
                        res.ContentType = Const.CONTENT_TYPE_TEXT;
                        byte[] text = Encoding.UTF8.GetBytes(Const.MSG_SUCCESS);
                        res.OutputStream.Write(text, 0, text.Length);

                        UpdateTextBoxLog("userInfo count=" + _userDatas.Count);

                        if (req.HttpMethod == Const.METHOD_POST)
                            UpdateTextBoxLog("post success");
                        else
                            UpdateTextBoxLog("put success");
                    }
                    else if (req.HttpMethod == Const.METHOD_GET || req.HttpMethod == Const.METHOD_DELETE)
                    {
                        if (!Utils.IsNumStr(reqSplit[2]))
                        {
                            // URLに正しいIDが指定されていない
                            res.StatusCode = 400;
                            res.StatusDescription = "ID is not number.";

                            UpdateTextBoxLog(res.StatusCode.ToString());
                            UpdateTextBoxLog(res.StatusDescription);
                            continue;
                        }

                        int id = int.Parse(reqSplit[2]);

                        if (req.HttpMethod == Const.METHOD_GET)
                        {
                            var retData = getUser(id);

                            if (retData == null)
                            {
                                // IDに紐づくデータが存在しない
                                res.StatusCode = 400;
                                res.StatusDescription = "ID is not exist.";

                                UpdateTextBoxLog(res.StatusCode.ToString());
                                UpdateTextBoxLog(res.StatusDescription);
                                continue;
                            }

                            res.ContentType = Const.CONTENT_TYPE_JSON;
                            res.ContentEncoding = Encoding.UTF8;

                            // JSONを返す
                            byte[] text = Encoding.UTF8.GetBytes(Utils.Serialize<UserData>(retData));
                            res.OutputStream.Write(text, 0, text.Length);

                            UpdateTextBoxLog("[RES]=" + Utils.ByteAryToStr(text));
                            UpdateTextBoxLog("userInfo count=" + _userDatas.Count);
                            UpdateTextBoxLog("get success");
                        }
                        else
                        {
                            // GETとは異なり、IDが存在しなくてもエラーにはしない

                            deleteUser(id);

                            res.ContentType = Const.CONTENT_TYPE_TEXT;
                            byte[] text = Encoding.UTF8.GetBytes(Const.MSG_SUCCESS);
                            res.OutputStream.Write(text, 0, text.Length);

                            UpdateTextBoxLog("userInfo count=" + _userDatas.Count);
                            UpdateTextBoxLog("delete success");
                        }
                    }
                    else
                    {
                        // 未サポートのメソッド
                        res.StatusCode = 405;
                        res.StatusDescription = "Not support method.";

                        UpdateTextBoxLog(res.StatusCode.ToString());
                        UpdateTextBoxLog(res.StatusDescription);
                    }
                }
                catch (Exception ex)
                {
                    res.StatusCode = 500;
                    res.StatusDescription = ex.Message;

                    UpdateTextBoxLog(res.StatusCode.ToString());
                    UpdateTextBoxLog(res.StatusDescription);
                }
                finally
                {
                    UpdateTextBoxLog("response close");
                    res.Close();
                }

            }// while
        }

        private void addUser(UserData data)
        {
            if (!_userDatas.ContainsKey(data.Id))
            {
                _userDatas.Add(data.Id, data.Name);
            }
        }

        private void updateUser(UserData data)
        {
            if (_userDatas.ContainsKey(data.Id))
            {
                _userDatas[data.Id] = data.Name;
            }
        }

        private UserData getUser(int id)
        {
            if (_userDatas.ContainsKey(id))
            {
                var ret = new UserData();
                ret.Id = id;
                ret.Name = _userDatas[id];
                return ret;
            }
            return null;
        }

        private void deleteUser(int id)
        {
            if (_userDatas.ContainsKey(id))
            {
                _userDatas.Remove(id);
            }
        }
    }
}
