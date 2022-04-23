using System;
using System.IO;
using System.Net;
using System.Net.Cache;
using System.Net.Security;
using System.Runtime.Serialization.Json;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using System.Windows.Forms;
using MyData;

namespace TinyHttpClient
{
    public partial class FormMain : Form
    {
        public FormMain()
        {
            InitializeComponent();
        }

        private void buttonPUT_Click(object sender, EventArgs e)
        {
            doUrl(Const.METHOD_PUT);
        }

        private void buttonDELETE_Click(object sender, EventArgs e)
        {
            doUrl(Const.METHOD_DELETE);
        }

        private void buttonGET_Click(object sender, EventArgs e)
        {
            doUrl(Const.METHOD_GET);
        }

        private void buttonPOST_Click(object sender, EventArgs e)
        {
            doUrl(Const.METHOD_POST);
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

        private void setTlsVersion(object tls)
        {
            try
            {
                // .NET3.5の場合はTLSv1がデフォルト（拡張するKBを当てないとv1.1以降は使えない）
                // つまり、.NET3.5だと以下のソースはビルドできない。

                if(tls == null || string.IsNullOrEmpty(tls.ToString()))
                {
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls;
                }
                else if(tls.ToString() == Const.TLS_11)
                {
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls11;
                }
                else if(tls.ToString() == Const.TLS_12)
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
            return true;  // SSL証明書の使用は問題なし
        }

        private void doUrl(string method)
        {
            try
            {
                var url = this.textBoxURL.Text;
                if(!Utils.IsUrl(url))
                {
                    MessageBox.Show("URL is bad!");
                    return;
                }

                var scheme = "http";
                if (radioButtonHttps.Checked)
                {
                    scheme = "https";
                }

                if (!url.StartsWith(scheme))
                {
                    MessageBox.Show("URL scheme is unmatch!");
                    return;
                }

                if (radioButtonHttps.Checked)
                {
                    setTlsVersion(this.comboBoxTLS.SelectedItem);

                    //自己証明書を突破するために以下を実施する
                    ServicePointManager.ServerCertificateValidationCallback =
                        new RemoteCertificateValidationCallback(OnRemoteCertificateValidationCallback);
                }

                var req = (HttpWebRequest)WebRequest.Create(url);
                req.Method = method;
                req.UserAgent = "testApp";
                req.ReadWriteTimeout = 5 * 1000;
                req.Timeout = 5 * 1000;
                req.CachePolicy = new RequestCachePolicy(RequestCacheLevel.NoCacheNoStore);
                req.KeepAlive = false;
                req.ContentType = "application/x-www-form-urlencoded";

                if(method == Const.METHOD_POST || method == Const.METHOD_PUT)
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

                var res = (HttpWebResponse)req.GetResponse();

                using(var s = res.GetResponseStream())
                {
                    using(var sr = new StreamReader(s, Encoding.UTF8))
                    {
                        MessageBox.Show(sr.ReadToEnd());
                    }
                }

                res.Close();
            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }


    }
}
