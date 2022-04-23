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
/// ・ステータスコードに対応しよう!
/// ・ログに対応しよう！
/// 
/// </summary>
namespace TinyRestServer
{
    public partial class Form1 : Form
    {
        private HttpListener listener = null;

        private Dictionary<int, String> userDatas = null;

        public Form1()
        {
            InitializeComponent();

            backgroundWorker1.WorkerReportsProgress = true;
            backgroundWorker1.WorkerSupportsCancellation = true;

            userDatas = new Dictionary<int, string>();
        }

        private void button_Start_Click(object sender, EventArgs e)
        {
            if (backgroundWorker1.IsBusy != true)
            {
                // Start the asynchronous operation.
                backgroundWorker1.RunWorkerAsync();
            }
        }

        private void button_Stop_Click(object sender, EventArgs e)
        {
            if (backgroundWorker1.WorkerSupportsCancellation == true)
            {
                // Cancel the asynchronous operation.
                backgroundWorker1.CancelAsync();
            }

            if (listener != null)
            {
                listener.Stop();
                listener.Close();
                listener = null;
            }
        }

        private void backgroundWorker1_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;

            listener = new HttpListener();
 
            // TODO:リッスンするホスト名とポートを指定します。
            //listener.Prefixes.Add("http://*:7016/");

            listener.Prefixes.Add(textBox_Url.Text + ":" + textBox_Port.Text + @"/");

            listener.Start();

            for (; ; )
            {
                if (worker.CancellationPending == true)
                {
                    e.Cancel = true;
                    break;
                }

                // リクエストが来るまでここで止まる
                var c = listener.GetContext();

                // リクエストを取得
                var req = c.Request;

                var reqSplit = req.RawUrl.Split('/');
                if (reqSplit.Length < 2)
                {
                    //TODO:エラー処理
                }

                var api = reqSplit[1];
                if (api != "users")
                {
                    //TODO:エラー処理
                }

                HttpListenerResponse res = null;

                if (req.HttpMethod == "POST" || req.HttpMethod == "PUT")
                {
                    if (req.HasEntityBody)
                    {
                        // リクエストボディに対する処理
                        using (var sr = new StreamReader(req.InputStream, new UTF8Encoding(false)))
                        {
                            var userData = Utils.Deserialize<UserData>(sr.ReadToEnd());

                            if (req.HttpMethod == "POST")
                            {
                                addUser(userData);
                            }
                            else
                            {
                                updateUser(userData);
                            }
                        }
                    }

                    try
                    {
                        res = c.Response;
                        res.ContentType = "text/plain";
                        new DataContractJsonSerializer(typeof(Result)).WriteObject(res.OutputStream, "success");
                    }
                    finally
                    {
                        res.Close();
                    }
                }
                else if (req.HttpMethod == "GET" || req.HttpMethod == "DELETE")
                {
                    int id = -1;

                    try
                    {
                        id = int.Parse(reqSplit[2]);
                    }
                    catch
                    {
                        //TODO:エラー処理
                    }

                    

                    if (req.HttpMethod == "GET")
                    {
                        var retData = getUser(id);

                        if (retData == null)
                        {
                            //TODO:エラー処理
                        }

                        try
                        {
                            res = c.Response;
                            res.ContentType = "application/json";
                            new DataContractJsonSerializer(typeof(Result)).WriteObject(res.OutputStream, Utils.Serialize<UserData>(retData));
                        }
                        finally
                        {
                            res.Close();
                        }
                    }
                    else
                    {
                        deleteUser(id);

                        try
                        {
                            res = c.Response;
                            res.ContentType = "text/plain";
                            new DataContractJsonSerializer(typeof(Result)).WriteObject(res.OutputStream, "success");
                        }
                        finally
                        {
                            res.Close();
                        }
                    }
                }
                else
                {
                    //TODO:エラー処理
                }
            }
        }

        private void addUser(UserData data)
        {
            if (!userDatas.ContainsKey(data.Id))
            {
                userDatas.Add(data.Id, data.Name);
            }
        }

        private void updateUser(UserData data)
        {
            if (userDatas.ContainsKey(data.Id))
            {
                userDatas[data.Id] = data.Name;
            }
        }

        private UserData getUser(int id)
        {
            if (userDatas.ContainsKey(id))
            {
                var ret = new UserData();
                ret.Id = id;
                ret.Name = userDatas[id];
                return ret;
            }
            return null;
        }

        private void deleteUser(int id)
        {
            if (userDatas.ContainsKey(id))
            {
                userDatas.Remove(id);
            }
        }
    }

    public class Result
    {
        public bool Value { get; set; }

        public string Message { get; set; }
    }
}
