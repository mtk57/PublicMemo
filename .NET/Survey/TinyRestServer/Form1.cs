using System;
using System.ComponentModel;
using System.IO;
using System.Net;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Windows.Forms;

/// <summary>
/// 簡易RESTサーバ
/// 
/// [memo]
/// ・exeを実行すると「アクセスが拒否されました」と出る場合は、マニフェストファイルのrequestedExecutionLevelのレベルを「requireAdministrator」に変更する。
///   参考：https://sweep3092.hatenablog.com/entry/2014/12/27/160005
/// 
/// [TODO]
/// MUST
/// ・Start押しても固まらないようにしよう!
/// ・Stopに対応しよう!
/// ・httpsに対応しよう!
/// ・Vistaで動くか試そう!
/// ・ステータスコードに対応しよう!
/// ・タイムアウトに対応しよう!
/// 
/// WANT
/// ・GET以外にも対応しよう!
/// ・動的にJSONを変更しよう!
/// 
/// </summary>
namespace TinyRestServer
{
    public partial class Form1 : Form
    {
        private HttpListener listener = null;

        public Form1()
        {
            InitializeComponent();

            backgroundWorker1.WorkerReportsProgress = true;
            backgroundWorker1.WorkerSupportsCancellation = true;
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

                var c = listener.GetContext();

                var req = c.Request;

                // TODO:ここでURLを切り分けます
                Console.WriteLine(req.RawUrl);

                if (req.HasEntityBody)
                {
                    // TODO：ここでリクエストボディに対する処理を行います。
                    using (var sr = new StreamReader(req.InputStream, new UTF8Encoding(false)))
                    {
                        Console.WriteLine(sr.ReadToEnd());
                    }
                }

                // JSONとして返したい値
                var value = new Result() { Value = true, Message = "あいうえおabcde" };

                // レスポンスにJSONを書き込みます。
                HttpListenerResponse res = null;
                try
                {
                    res = c.Response;
                    res.ContentType = "application/json";
                    new DataContractJsonSerializer(typeof(Result)).WriteObject(res.OutputStream, value);
                }
                finally
                {
                    res.Close();

                }
            }
            
        }
    }

    public class Result
    {
        public bool Value { get; set; }

        public string Message { get; set; }
    }
}
