using System;
using System.IO;
using System.Net;
using System.Net.Cache;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using CommonLib;
using System.Collections.Generic;

// Create : POST
// Read	  : GET
// Update : PUT
// Delete : DELETE

namespace nsLibCOM
{
	[ComVisible(true)]
	public interface ILibCOM
	{
		void Init();
		void Dispose();
		void ReloadSettings();
		string PostToken();
	}

	[ClassInterface(ClassInterfaceType.None)]
	public class LibCOM: ILibCOM
	{
		private const string SETTINGS_JSON = "Settings.json";
		private const string LOG = "LibCOM.log";
		private const string CONTENT_TYPE = "application/x-www-form-urlencoded";
		private const string CONTENT_TYPE_JSON = "application/json";

		// ★本番環境に合わせること
		private const string KEY_1 = "client_id";
		private const string KEY_2 = "client_sec";
		private const string KEY_3 = "user_name";
		private const string KEY_4 = "user_pw";
		private const string KEY_5 = "g_type";

		private string _settingsJsonPath = "";
		private Settings _settings;

		public void Init()
		{
			Logger.Initialize(getLogPath());
			Logger.Write("Init S");
			ReloadSettings();
			Logger.Write("Init E");
		}

		public void Dispose()
		{
			Logger.Write("Dispose S");
			Logger.Write("Dispose E");
			Logger.Dispose();
		}

		public void ReloadSettings()
		{
			var myDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
			_settingsJsonPath = Path.Combine(myDir, SETTINGS_JSON);
			if (!File.Exists(_settingsJsonPath))
			{
				throw new Exception("Settings.json is not exist! (" + _settingsJsonPath + ")");
			}
			_settings = getSettings();
		}

		public string PostToken()
		{
			Logger.Write("PostToken S");

			var settings = HttpSettings.ConvertFromPostToken(_settings);
			
			var req = request(settings);

			var token = response(req, settings) as string;

			Logger.Write("token=" + token);
			Logger.Write("PostToken E");

			return token;
		}

		private string getLogPath()
		{
			var myDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
			return Path.Combine(myDir, LOG);
		}	

		private Settings getSettings()
		{
			var json = File.ReadAllText(_settingsJsonPath);
			return SettingUtil.Deserialize<Settings>(json);
		}

		private void setSettings(Settings settings)
		{
			var serialize = SettingUtil.Serialize<Settings>(settings);
			File.WriteAllText(_settingsJsonPath, serialize);
		}

		private HttpWebRequest request(HttpSettings settings)
		{
			Logger.Write("request S");

			var req = (HttpWebRequest)WebRequest.Create(settings.url);
			req.UserAgent = "LibCOM";
			req.ReadWriteTimeout = settings.rw_timeout_sec * 1000;
			req.Timeout = settings.timeout_sec * 1000;
			req.CachePolicy = new RequestCachePolicy(RequestCacheLevel.NoCacheNoStore);
			req.KeepAlive = false;

			req.Method = settings.method;
			req.ContentType = settings.content_type;

			if (settings.method == "POST" || settings.method == "PUT")
			{
				byte[] data;

				if (settings.content_type == "application/x-www-form-urlencoded")
				{
					data = Encoding.ASCII.GetBytes(settings.data);
					req.ContentLength = data.Length;

					using (var s = req.GetRequestStream())
					{
						s.Write(data, 0, data.Length);
					}
				}
				else if (settings.content_type == "application/json")
				{
					// TODO
				}
				else
				{
					// 未サポート
					throw new Exception("Not support content type! (" + settings.content_type + ")");
				}
			}
			else if (settings.method == "GET" || settings.method == "DELETE")
			{
				// TODO:
			}
			else
			{
				// 未サポート
				throw new Exception("Not support method! (" + settings.method + ")");
			}

			Logger.Write("request E");

			return req;
		}

		private object response(HttpWebRequest req, HttpSettings settings)
		{
			Logger.Write("response S");

			object ret = null;
			HttpWebResponse res = null;

			try
			{
				res = (HttpWebResponse)req.GetResponse();

				Logger.Write("HttpStatus=" + res.StatusCode);
				Logger.Write("StatusDescription=" + res.StatusDescription);

				using (var s = res.GetResponseStream())
				{
					using (var sr = new StreamReader(s, Encoding.UTF8))
					{
						ret = sr.ReadToEnd();

						if (settings.IsJson)
						{
							// TODO:
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

				Logger.Write(msg);

				throw new Exception("WebException! Check the log file.");
			}
			finally
			{
				if (res != null)
				{
					res.Close();
				}
			}

			Logger.Write("ret=" + ret);
			Logger.Write("response E");

			return ret;
		}

		private class HttpSettings
		{
			public int rw_timeout_sec { get; set; }
			public int timeout_sec { get; set; }
			public string method { get; set; }
			public string url { get; set; }
			public string content_type { get; set; }
			public string data { get; set; }

			public bool IsJson
			{
				get
				{
					return content_type == CONTENT_TYPE_JSON;
				}
			}

			public static HttpSettings ConvertFromPostToken(Settings settings)
			{
				var ret = new HttpSettings();

				ret.rw_timeout_sec = settings.post_token.rw_timeout_sec;
				ret.timeout_sec = settings.post_token.timeout_sec;
				ret.method = settings.post_token.method;
				ret.url = settings.post_token.url;
				ret.content_type = settings.post_token.content_type;

				var dict = new Dictionary<string, string>();

				if (!string.IsNullOrEmpty(settings.post_token.client_id))
				{
					dict[KEY_1] = settings.post_token.client_id;
				}
				if (!string.IsNullOrEmpty(settings.post_token.client_sec))
				{
					dict[KEY_2] = settings.post_token.client_sec;
				}
				if (!string.IsNullOrEmpty(settings.post_token.user_name))
				{
					dict[KEY_3] = settings.post_token.user_name;
				}
				if (!string.IsNullOrEmpty(settings.post_token.user_pw))
				{
					dict[KEY_4] = settings.post_token.user_pw;
				}
				if (!string.IsNullOrEmpty(settings.post_token.g_type))
				{
					dict[KEY_5] = settings.post_token.g_type;
				}

				foreach (var k in dict.Keys)
				{
					ret.data += String.Format("{0}={1}&", k, dict[k]);
				}

				return ret;
			}
		}
	}

}