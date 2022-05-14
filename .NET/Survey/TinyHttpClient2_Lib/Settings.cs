using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;

namespace nsLibCOM
{
    public static class SettingUtil
    {
        public static T Deserialize<T>(string json)
        {
            T result;
            var serializer = new DataContractJsonSerializer(typeof(T));

            using (var ms = new MemoryStream(Encoding.UTF8.GetBytes(json)))
            {
                result = (T)serializer.ReadObject(ms);
            }
            return result;
        }

        public static string Serialize<T>(T obj)
        {
            using (var ms = new MemoryStream())
            {
                var serializer = new DataContractJsonSerializer(typeof(T));
                serializer.WriteObject(ms, obj);
                return Encoding.UTF8.GetString(ms.ToArray());
            }
        }
    }

    [DataContract]
    public class Settings
    {
        [DataMember(Name = "post_token")]
        public PostToken post_token { get; set; }
    }

    [DataContract]
    public class PostToken
    {
        [DataMember(Name = "rw_timeout_sec")]
        public int rw_timeout_sec { get; set; }

        [DataMember(Name = "timeout_sec")]
        public int timeout_sec { get; set; }

        [DataMember(Name = "method")]
        public string method { get; set; }

        [DataMember(Name = "url")]
        public string url { get; set; }

        [DataMember(Name = "content_type")]
        public string content_type { get; set; }

        [DataMember(Name = "client_id")]
        public string client_id { get; set; }

        [DataMember(Name = "user_name")]
        public string user_name { get; set; }

        [DataMember(Name = "user_pw")]
        public string user_pw { get; set; }

        [DataMember(Name = "client_sec")]
        public string client_sec { get; set; }

        [DataMember(Name = "g_type")]
        public string g_type { get; set; }
    }
}
