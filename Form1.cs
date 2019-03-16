using System;
using System.IO;
using System.Net;
using System.Net.Cache;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Web;
using System.Windows.Forms;

namespace TestHttps
{
    public partial class FormMain : Form
    {
        public FormMain()
        {
            InitializeComponent();
        }

        private void buttonGET_Click(object sender, EventArgs e)
        {
            doUrl("GET");
        }

        private void buttonPOST_Click(object sender, EventArgs e)
        {
            doUrl("POST");
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
                else if(tls.ToString() == "v1.1")
                {
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls11;
                }
                else if(tls.ToString() == "v1.2")
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
                string url = this.textBoxURL.Text;
                if(url.Length < 10)
                {
                    MessageBox.Show("URLを入力してください。");
                    return;
                }

                setTlsVersion(this.comboBoxTLS.SelectedItem);

                // XAMPPのApacheはオレオレ証明書なので突破するために以下を実施
                ServicePointManager.ServerCertificateValidationCallback =
                    new RemoteCertificateValidationCallback(OnRemoteCertificateValidationCallback);

                string param = "";
                string prm1 = this.textBoxParam1.Text;
                string prm2 = this.textBoxParam2.Text;
                string val1 = this.textBoxValue1.Text;
                string val2 = this.textBoxValue1.Text;

                Encoding enc = Encoding.UTF8;

                if(prm1.Length > 1)
                {
                    param = prm1 + "=" + HttpUtility.UrlEncode(val1, enc);
                }

                if(prm2.Length > 1)
                {
                    if(param.Length > 0)
                    {
                        param += "&";
                    }
                    param += prm2 + "=" + HttpUtility.UrlEncode(val2, enc);
                }

                if(method == "GET")
                {
                    if(param.Length > 0)
                    {
                        url = url + "?" + param;
                    }
                }

                HttpWebRequest webReq = (HttpWebRequest)WebRequest.Create(url);
                webReq.Method = method;
                webReq.UserAgent = "testApp";
                webReq.ReadWriteTimeout = 5 * 1000;
                webReq.Timeout = 5 * 1000;
                webReq.CachePolicy = new RequestCachePolicy(RequestCacheLevel.NoCacheNoStore);
                //webReq.KeepAlive = true;
                webReq.KeepAlive = false;
                //webReq.Proxy = webReq.GetDefaultProxy();

                webReq.ContentType = "application/x-www-form-urlencoded";
                if(method == "POST")
                {
                    webReq.ContentLength = param.Length;
                    using(Stream paraStream = webReq.GetRequestStream())
                    {
                        using(StreamWriter sw = new StreamWriter(paraStream))
                        {
                            sw.Write(param);
                        }
                    }
                }

                HttpWebResponse webRes = (HttpWebResponse)webReq.GetResponse();

                using(Stream resStream = webRes.GetResponseStream())
                {
                    using(StreamReader sr = new StreamReader(resStream, enc))
                    {
                        MessageBox.Show(sr.ReadToEnd());
                    }
                }

                webRes.Close();
            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }
    }
}
