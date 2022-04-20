using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Text.Json;
using System.Windows.Forms;
using Newtonsoft.Json;

//
// JSONデータのデシリアライズ/シリアライズのテスト
//
//●使用したライブラリ
//  [.NET標準]
//  ・System.Runtime.Serialization.Json.DataContractJsonSerializer  (.NET3.5～)
//  ・System.Text.Json.JsonSerializer                               (.NET4.6.1～)
//  [サードパーティ]
//  ・Newtonsoft.Json.JsonConvert                                   (JSON.NET)
//
//●型
//  ・文字列("...")
//  ・数値(123, 12.3, 1.23e4 など)
//  ・ヌル値(null)
//  ・真偽値(true, false)
//  ・オブジェクト({ ... })
//  ・配列([...])
//
//●エンコーディング
//  ・RFC 8259 で指定された仕様では、BOM 無しの UTF-8 で記述する(MUST)と定義されています。
//
//●注意点
//  ・JsonSerializer以外は、DataMember属性のNameが効いたので、メンバ変数名とJSONの要素名が一致させなくてもよい。
//    JsonSerializerは効かないので、メンバ変数名とJSONの要素名を一致させる必要がある。
//
namespace WindowsFormsApp
{
    public partial class Form1 : Form
    {
        #region テスト用のJSONデータ

        // 文字列
        private string JSON_STR = @"{
""name"": ""Tanaka""
}";

        // 数値
        private string JSON_NUM = @"{
""age"": 26,
""pi"": 3.14,
""planck_constant"": 6.62607e-34
}";

        // ヌル値
        private string JSON_NULL = @"{
""name"": null
}";

        // 真偽値
        private string JSON_BOOL = @"{
 ""active_flag"": true,
 ""delete_flag"": false
}";

        // オブジェクト
        private string JSON_OBJ = @"{
 ""user_info"": {
 ""user_id"": ""A1234567"",
 ""user_name"": ""Yamada Taro""
 }
 }";

        // 配列（オブジェクト内）
        private string JSON_ARY = @"{
""color_list"": [ ""red"", ""green"", ""blue"" ],
""num_list"": [ 123, 456, 789 ],
""mix_list"": [ ""red"", 456, null, true ],
""array_list"": [ [ 12, 23 ], [ 34, 45 ], [ 56, 67 ] ],
""object_list"": [
{ ""name"": ""Tanaka"", ""age"": 26 },
{ ""name"": ""Suzuki"", ""age"": 32 }
]
}";

        // 配列（ルート）
        private string JSON_ARY_ROOT = @"[
{
""name"": ""Taro"",
""age"": 30,
""languages"": [""Japanese"", ""English""],
""active"": true
},
{
""name"": ""Aiko"",
""age"": 33,
""languages"": [""Japanese""],
""active"": false
},
{
""name"": ""Hanako"",
""age"": 29,
""languages"": [""English"", ""French""],
""active"": true
}
]";

        #endregion テスト用のJSONデータ

        private enum LibType
        {
            DataContractJsonSerializer,
            JsonSerializer,
            JsonDotNet
        }


        public Form1()
        {
            InitializeComponent();
        }


        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                DeserializeSerializer_Test(LibType.DataContractJsonSerializer);

                MessageBox.Show("Success!");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                DeserializeSerializer_Test(LibType.JsonSerializer);

                MessageBox.Show("Success!");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                DeserializeSerializer_Test(LibType.JsonDotNet);

                MessageBox.Show("Success!");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void DeserializeSerializer_Test(LibType type)
        {
            string serialize = null;

            #region 文字列
            DataString data_string = null;
            if (type == LibType.DataContractJsonSerializer)
            {
                data_string = Deserialize_By_DataContractJsonSerializer<DataString>(JSON_STR);
            }
            else if (type == LibType.JsonSerializer)
            {
                data_string = Deserialize_By_JsonSerializer<DataString>(JSON_STR);
            }
            else if (type == LibType.JsonDotNet)
            {
                data_string = Deserialize_By_JsonDotNet<DataString>(JSON_STR);
            }
            Console.WriteLine($"IN={JSON_STR}");
            Console.WriteLine($"OUT=");
            Console.WriteLine($"{data_string.name}");
            Console.WriteLine("------------");
            if (type == LibType.DataContractJsonSerializer)
            {
                serialize = Serialize_By_DataContractJsonSerializer<DataString>(data_string);
            }
            else if (type == LibType.JsonSerializer)
            {
                serialize = Serialize_By_JsonSerializer(data_string);
            }
            else if (type == LibType.JsonDotNet)
            {
                serialize = Serialize_By_JsonDotNet(data_string);
            }
            Console.WriteLine($"serialize={serialize}");
            Console.WriteLine("============");
            #endregion 文字列

            #region 数値
            DataNum data_num = null;
            if (type == LibType.DataContractJsonSerializer)
            {
                data_num = Deserialize_By_DataContractJsonSerializer<DataNum>(JSON_NUM);
            }
            else if (type == LibType.JsonSerializer)
            {
                data_num = Deserialize_By_JsonSerializer<DataNum>(JSON_NUM);
            }
            else if (type == LibType.JsonDotNet)
            {
                data_num = Deserialize_By_JsonDotNet<DataNum>(JSON_NUM);
            }
            Console.WriteLine($"IN={JSON_NUM}");
            Console.WriteLine($"OUT=");
            Console.WriteLine($"{data_num.age}");
            Console.WriteLine($"{data_num.pi}");
            Console.WriteLine($"{data_num.planck_constant}");
            Console.WriteLine("------------");
            if (type == LibType.DataContractJsonSerializer)
            {
                serialize = Serialize_By_DataContractJsonSerializer<DataNum>(data_num);
            }
            else if (type == LibType.JsonSerializer)
            {
                serialize = Serialize_By_JsonSerializer(data_num);
            }
            else if (type == LibType.JsonDotNet)
            {
                serialize = Serialize_By_JsonDotNet(data_num);
            }
            Console.WriteLine($"serialize={serialize}");
            Console.WriteLine("============");
            #endregion 数値

            #region ヌル値
            DataNull data_null = null;
            if (type == LibType.DataContractJsonSerializer)
            {
                data_null = Deserialize_By_DataContractJsonSerializer<DataNull>(JSON_NULL);
            }
            else if (type == LibType.JsonSerializer)
            {
                data_null = Deserialize_By_JsonSerializer<DataNull>(JSON_NULL);
            }
            else if (type == LibType.JsonDotNet)
            {
                data_null = Deserialize_By_JsonDotNet<DataNull>(JSON_NULL);
            }
            Console.WriteLine($"IN={JSON_NULL}");
            Console.WriteLine($"OUT=");
            Console.WriteLine($"{data_null.name}");
            Console.WriteLine("------------");
            if (type == LibType.DataContractJsonSerializer)
            {
                serialize = Serialize_By_DataContractJsonSerializer<DataNull>(data_null);
            }
            else if (type == LibType.JsonSerializer)
            {
                serialize = Serialize_By_JsonSerializer(data_null);
            }
            else if (type == LibType.JsonDotNet)
            {
                serialize = Serialize_By_JsonDotNet(data_null);
            }
            Console.WriteLine($"serialize={serialize}");
            Console.WriteLine("============");
            #endregion ヌル値

            #region 真偽値
            DataBool data_bool = null;
            if (type == LibType.DataContractJsonSerializer)
            {
                data_bool = Deserialize_By_DataContractJsonSerializer<DataBool>(JSON_BOOL);
            }
            else if (type == LibType.JsonSerializer)
            {
                data_bool = Deserialize_By_JsonSerializer<DataBool>(JSON_BOOL);
            }
            else if (type == LibType.JsonDotNet)
            {
                data_bool = Deserialize_By_JsonDotNet<DataBool>(JSON_BOOL);
            }
            Console.WriteLine($"IN={JSON_BOOL}");
            Console.WriteLine($"OUT=");
            Console.WriteLine($"{data_bool.active_flag}");
            Console.WriteLine($"{data_bool.delete_flag}");
            Console.WriteLine("------------");
            if (type == LibType.DataContractJsonSerializer)
            {
                serialize = Serialize_By_DataContractJsonSerializer<DataBool>(data_bool);
            }
            else if (type == LibType.JsonSerializer)
            {
                serialize = Serialize_By_JsonSerializer(data_bool);
            }
            else if (type == LibType.JsonDotNet)
            {
                serialize = Serialize_By_JsonDotNet(data_bool);
            }
            Console.WriteLine($"serialize={serialize}");
            Console.WriteLine("============");
            #endregion 真偽値

            #region オブジェクト
            DataObj data_obj = null;
            if (type == LibType.DataContractJsonSerializer)
            {
                data_obj = Deserialize_By_DataContractJsonSerializer<DataObj>(JSON_OBJ);
            }
            else if (type == LibType.JsonSerializer)
            {
                data_obj = Deserialize_By_JsonSerializer<DataObj>(JSON_OBJ);
            }
            else if (type == LibType.JsonDotNet)
            {
                data_obj = Deserialize_By_JsonDotNet<DataObj>(JSON_OBJ);
            }
            Console.WriteLine($"IN={JSON_OBJ}");
            Console.WriteLine($"OUT=");
            Console.WriteLine($"{data_obj.user_info.user_id}");
            Console.WriteLine($"{data_obj.user_info.user_name}");
            Console.WriteLine("------------");
            if (type == LibType.DataContractJsonSerializer)
            {
                serialize = Serialize_By_DataContractJsonSerializer<DataObj>(data_obj);
            }
            else if (type == LibType.JsonSerializer)
            {
                serialize = Serialize_By_JsonSerializer(data_obj);
            }
            else if (type == LibType.JsonDotNet)
            {
                serialize = Serialize_By_JsonDotNet(data_obj);
            }
            Console.WriteLine($"serialize={serialize}");
            Console.WriteLine("============");
            #endregion オブジェクト

            #region 配列（オブジェクト内）
            DataAry data_ary = null;
            if (type == LibType.DataContractJsonSerializer)
            {
                data_ary = Deserialize_By_DataContractJsonSerializer<DataAry>(JSON_ARY);
            }
            else if (type == LibType.JsonSerializer)
            {
                data_ary = Deserialize_By_JsonSerializer<DataAry>(JSON_ARY);
            }
            else if (type == LibType.JsonDotNet)
            {
                data_ary = Deserialize_By_JsonDotNet<DataAry>(JSON_ARY);
            }
            Console.WriteLine($"IN={JSON_ARY}");
            Console.WriteLine($"OUT=");
            Console.WriteLine(string.Join(", ", data_ary.color_list));
            Console.WriteLine(string.Join(", ", data_ary.num_list));
            Console.WriteLine(string.Join(", ", data_ary.mix_list));
            foreach (var ary in data_ary.array_list)
            {
                Console.WriteLine(string.Join(", ", ary));
            }
            foreach (var obj in data_ary.object_list)
            {
                Console.WriteLine($"{obj.name}, {obj.age}");
            }
            Console.WriteLine("------------");
            if (type == LibType.DataContractJsonSerializer)
            {
                serialize = Serialize_By_DataContractJsonSerializer<DataAry>(data_ary);
            }
            else if (type == LibType.JsonSerializer)
            {
                serialize = Serialize_By_JsonSerializer(data_ary);
            }
            else if (type == LibType.JsonDotNet)
            {
                serialize = Serialize_By_JsonDotNet(data_ary);
            }
            Console.WriteLine($"serialize={serialize}");
            Console.WriteLine("============");
            #endregion 配列（オブジェクト内）


            #region 配列（ルート）
            List<UserInfoEx> data_objs = null;
            if (type == LibType.DataContractJsonSerializer)
            {
                data_objs = Deserialize_By_DataContractJsonSerializer<List<UserInfoEx>>(JSON_ARY_ROOT);
            }
            else if (type == LibType.JsonSerializer)
            {
                data_objs = Deserialize_By_JsonSerializer<List<UserInfoEx>>(JSON_ARY_ROOT);
            }
            else if (type == LibType.JsonDotNet)
            {
                data_objs = Deserialize_By_JsonDotNet<List<UserInfoEx>>(JSON_ARY_ROOT);
            }
            Console.WriteLine($"IN={JSON_ARY_ROOT}");
            Console.WriteLine($"OUT=");
            foreach (var obj in data_objs)
            {
                Console.WriteLine($"{obj.name}");
                Console.WriteLine($"{obj.age}");
                Console.WriteLine($"{string.Join(", ", obj.languages)}");
                Console.WriteLine($"{obj.active}");
            }
            Console.WriteLine("------------");
            if (type == LibType.DataContractJsonSerializer)
            {
                serialize = Serialize_By_DataContractJsonSerializer<List<UserInfoEx>>(data_objs);
            }
            else if (type == LibType.JsonSerializer)
            {
                serialize = Serialize_By_JsonSerializer(data_objs);
            }
            else if (type == LibType.JsonDotNet)
            {
                serialize = Serialize_By_JsonDotNet(data_objs);
            }
            Console.WriteLine($"serialize={serialize}");
            Console.WriteLine("============");
            #endregion 配列（ルート）

        }

        private T Deserialize_By_DataContractJsonSerializer<T>(string json)
        {
            T result;
            var serializer = new DataContractJsonSerializer(typeof(T));

            using (var ms = new MemoryStream(Encoding.UTF8.GetBytes(json)))
            {
                result = (T)serializer.ReadObject(ms);
            }
            return result;
        }
        private string Serialize_By_DataContractJsonSerializer<T>(T obj)
        {
            using (var ms = new MemoryStream())
            {
                var serializer = new DataContractJsonSerializer(typeof(T));
                serializer.WriteObject(ms, obj);
                return Encoding.UTF8.GetString(ms.ToArray());
            }
        }

        private T Deserialize_By_JsonSerializer<T>(string json)
        {
            return System.Text.Json.JsonSerializer.Deserialize<T>(json);
        }
        private string Serialize_By_JsonSerializer<T>(T obj)
        {
            return System.Text.Json.JsonSerializer.Serialize(obj);
        }

        private T Deserialize_By_JsonDotNet<T>(string json)
        {
            return JsonConvert.DeserializeObject<T>(json);
        }

        private string Serialize_By_JsonDotNet<T>(T obj)
        {
            return JsonConvert.SerializeObject(obj);
        }


        #region 文字列
        [DataContract]
        class DataString
        {
            [DataMember(Name = "name")]
            public string name { get; set; }
        }
        #endregion 文字列

        #region 数値
        [DataContract]
        class DataNum
        {
            [DataMember(Name = "age")]
            public int age { get; set; }

            [DataMember(Name = "pi")]
            public double pi { get; set; }

            [DataMember(Name = "planck_constant")]
            public double planck_constant { get; set; }
        }
        #endregion 数値

        #region ヌル値
        [DataContract]
        class DataNull
        {
            [DataMember(Name = "name")]
            public string name { get; set; }
        }
        #endregion ヌル値

        #region 真偽値
        [DataContract]
        class DataBool
        {
            [DataMember(Name = "active_flag")]
            public bool active_flag { get; set; }

            [DataMember(Name = "delete_flag")]
            public bool delete_flag { get; set; }
        }
        #endregion 真偽値

        #region オブジェクト
        [DataContract]
        class DataObj
        {
            [DataMember(Name = "user_info")]
            public UserInfo user_info { get; set; }
        }

        [DataContract]
        class UserInfo
        {
            [DataMember(Name = "user_id")]
            public string user_id { get; set; }

            [DataMember(Name = "user_name")]
            public string user_name { get; set; }
        }

        [DataContract]
        class UserInfoEx
        {
            [DataMember(Name = "name")]
            public string name { get; set; }

            [DataMember(Name = "age")]
            public int age { get; set; }

            [DataMember(Name = "languages")]
            public List<string> languages { get; set; }

            [DataMember(Name = "active")]
            public bool active { get; set; }
        }
        #endregion オブジェクト

        #region 配列
        [DataContract]
        class DataAry
        {
            [DataMember(Name = "color_list")]
            public List<string> color_list { get; set; }

            [DataMember(Name = "num_list")]
            public List<int> num_list { get; set; }

            [DataMember(Name = "mix_list")]
            public List<object> mix_list { get; set; }

            [DataMember(Name = "array_list")]
            public List<List<int>> array_list { get; set; }

            [DataMember(Name = "object_list")]
            public List<PersonInfo> object_list { get; set; }
        }

        [DataContract]
        class PersonInfo
        {
            [DataMember(Name = "name")]
            public string name { get; set; }

            [DataMember(Name = "age")]
            public int age { get; set; }
        }

        #endregion 配列


    }
}
