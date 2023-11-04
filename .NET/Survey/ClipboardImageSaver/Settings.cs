using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;

namespace ClipboardImageSaver
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
        [DataMember(Name = "SaveDirPath")]
        public string SaveDirPath { get; set; }

        [DataMember(Name = "SaveFileName")]
        public string SaveFileName { get; set; }

        [DataMember(Name = "StartNum")]
        public decimal StartNum { get; set; }

        public override string ToString()
        {
            return $"SaveDirPath={SaveDirPath}, SaveFileName={SaveFileName}, StartNum={StartNum}";
        }
    }
}
