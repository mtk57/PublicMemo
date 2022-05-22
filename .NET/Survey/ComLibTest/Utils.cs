using System.IO;
using System.Reflection;
using System.Runtime.Serialization.Json;
using System.Text;

namespace NsComLibTest
{
    internal class Utils
    {
        public static string GetResDir()
        {
            return GetMyDir() + @"\res";
        }

        public static string GetMyDir()
        {
            return Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        }

        public static T ReadJsonFile<T>(string path)
        {
            return Deserialize<T>(File.ReadAllText(path));
        }

        public static T Deserialize<T>(string json)
        {
            T result;
            var serializer = new DataContractJsonSerializer(typeof(T));
            using(var ms = new MemoryStream(Encoding.UTF8.GetBytes(json)))
            {
                result = (T)serializer.ReadObject(ms);
            }
            return result;
        }

        public static string Serialize<T>(T obj)
        {
            using(var ms = new MemoryStream())
            {
                var serializer = new DataContractJsonSerializer(typeof(T));
                serializer.WriteObject(ms, obj);
                return Encoding.UTF8.GetString(ms.ToArray());
            }
        }
    }
}
