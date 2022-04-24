using MyData;
using System;
using System.IO;
using System.Net;
using System.Net.Cache;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Windows.Forms;

namespace TinyHttpClient
{
    public partial class FormMain : Form
    {
        private string _method = null;
        private string _url = null;

        public FormMain()
        {
            InitializeComponent();

            comboBoxTLS.Items.AddRange(new string[] { Const.TLS_10, Const.TLS_11, Const.TLS_12 });
        }

        private void buttonPOST_Click(object sender, EventArgs e)
        {
            executeWebApi(Const.METHOD_POST);
        }

        private void buttonGET_Click(object sender, EventArgs e)
        {
            executeWebApi(Const.METHOD_GET);
        }

        private void buttonPUT_Click(object sender, EventArgs e)
        {
            executeWebApi(Const.METHOD_PUT);
        }

        private void buttonDELETE_Click(object sender, EventArgs e)
        {
            executeWebApi(Const.METHOD_DELETE);
        }

        private void buttonDefaultURL_Click(object sender, EventArgs e)
        {
            if (radioButtonHttp.Checked)
            {
                textBoxURL.Text = Const.DEFAULT_URL_HTTP;
            }
            else
            {
                textBoxURL.Text = Const.DEFAULT_URL_HTTPS;
            }
        }

        private void buttonDefaultParam_Click(object sender, EventArgs e)
        {
            textBoxValueId.Text = Const.DEFAULT_ID.ToString();
            textBoxValueName.Text = Const.DEFAULT_NAME;
        }

        private void updateTextBoxLog(string data)
        {
            textBox_Log.AppendText(data + Environment.NewLine);
        }

        private void executeWebApi(string method)
        {
            _method = method;
            _url = textBoxURL.Text;

            if (!checkParama())
            {
                return;
            }

            updateTextBoxLog(_method);
            updateTextBoxLog(_url);

            if (radioButtonHttps.Checked)
            {
                setTlsVersion(comboBoxTLS.SelectedItem);

                //自己証明書を突破するために以下を実施する
                ServicePointManager.ServerCertificateValidationCallback =
                    new RemoteCertificateValidationCallback(OnRemoteCertificateValidationCallback);
            }

            try
            {
                updateTextBoxLog("[REQ]Start");

                HttpWebRequest req = request();

                updateTextBoxLog("[REQ]End");

                updateTextBoxLog("[RES]Start");

                response(req);

                updateTextBoxLog("[RES]End");
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                updateTextBoxLog(e.Message);
            }
            finally
            {
                updateTextBoxLog("-------------------");
            }
        }

        private bool checkParama()
        {
            if (!Utils.IsUrl(_url))
            {
                var msg = string.Format("URL is bad! [{0}]", _url);
                MessageBox.Show(msg);
                updateTextBoxLog(msg);
                return false;
            }

            var scheme = "http";
            if (radioButtonHttps.Checked)
            {
                scheme = "https";
            }

            if (!_url.StartsWith(scheme))
            {
                var msg = string.Format("URL scheme is unmatch! [{0}]", _url);
                MessageBox.Show(msg);
                updateTextBoxLog(msg);
                return false;
            }

            if (_method == Const.METHOD_POST || _method == Const.METHOD_PUT)
            {
                if (!Utils.IsNumStr(textBoxValueId.Text))
                {
                    var msg = string.Format("ID is not number! [{0}]", textBoxValueId.Text);
                    MessageBox.Show(msg);
                    updateTextBoxLog(msg);
                    return false;
                }
            }

            return true;
        }

        private void setTlsVersion(object tls)
        {
            try
            {
                // .NET3.5の場合はTLSv1がデフォルト（拡張するKBを当てないとv1.1以降は使えない）
                // つまり、.NET3.5だと以下のソースはビルドできない。

                if (tls == null || string.IsNullOrEmpty(tls.ToString()))
                {
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls;
                }
                else if (tls.ToString() == Const.TLS_11)
                {
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls11;
                }
                else if (tls.ToString() == Const.TLS_12)
                {
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                }
                else
                {
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls;
                }
            }
            catch
            {
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls;
            }
        }

        // 信頼できないSSL証明書を「問題なし」にする
        private bool OnRemoteCertificateValidationCallback(
            Object sender,
            X509Certificate certificate,
            X509Chain chain,
            SslPolicyErrors sslPolicyErrors)
        {
            // SSL証明書の使用は問題なし
            return true;
        }

        private HttpWebRequest request()
        {
            var req = (HttpWebRequest)WebRequest.Create(_url);
            req.Method = _method;
            req.UserAgent = "TinyHttpClient";
            req.ReadWriteTimeout = 30 * 1000;
            req.Timeout = 180 * 1000;
            req.CachePolicy = new RequestCachePolicy(RequestCacheLevel.NoCacheNoStore);
            req.KeepAlive = false;
            req.ContentType = "application/x-www-form-urlencoded";

            if (_method == Const.METHOD_POST || _method == Const.METHOD_PUT)
            {
                req.Accept = "application/json";
                req.ContentType = "application/json;";
                //req.ContentLength = param.Length;

                using (var s = req.GetRequestStream())
                {
                    var userData = new UserData();
                    userData.Id = int.Parse(textBoxValueId.Text);
                    userData.Name = textBoxValueName.Text;

                    var json = Utils.Serialize<UserData>(userData);
                    using (var sw = new StreamWriter(s))
                    {
                        sw.Write(json);
                    }
                }
            }
            return req;
        }

        private void response(HttpWebRequest req)
        {
            HttpWebResponse res = null;

            try
            {
                res = (HttpWebResponse)req.GetResponse();

                using (var s = res.GetResponseStream())
                {
                    using (var sr = new StreamReader(s, Encoding.UTF8))
                    {
                        var resData = sr.ReadToEnd();

                        updateTextBoxLog("[RES Data] " + resData);

                        if (_method == Const.METHOD_GET)
                        {
                            var userData = Utils.Deserialize<UserData>(resData);

                            updateTextBoxLog(string.Format("[Deserialize] id={0}, name={1}", userData.Id, userData.Name));
                        }
                    }
                }
            }
            catch (WebException e)
            {
                string msg;

                if (e.Status == WebExceptionStatus.ProtocolError)
                {
                    var errres = (HttpWebResponse)e.Response;
                    msg = string.Format("ERROR!\n{0}({1})\n{2}",
                            errres.StatusCode, (int)errres.StatusCode, errres.StatusDescription);

                }
                else
                {
                    msg = string.Format("ERROR!\n{0}\n{1}",
                            e.Status, e.Message);
                }
                MessageBox.Show(msg);
                updateTextBoxLog(msg);
            }
            finally
            {
                if(res != null)
                {
                    res.Close();
                }
            }
        }
    }
}
