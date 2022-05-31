using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;

namespace NsComLibTest
{
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IComLibTest
    {
        UserInfos GetUserInfos();

        string StructArgsTest(ref StructTest stTest);

        string SimpleFunc();
    }

    public struct StructTest
    {
        [MarshalAs(UnmanagedType.BStr)]
        public string x;
        public int y;
    }

    [ClassInterface(ClassInterfaceType.None)]
    public class ComLibMain : IComLibTest
    {
        public UserInfos GetUserInfos()
        {
            var userInfos = new UserInfos();
            var path = Path.Combine(Utils.GetResDir(), "UserInfos.json");
            userInfos.UserInfoList = Utils.ReadJsonFile<UserInfo[]>(path);

            return userInfos;
        }

        public string StructArgsTest(ref StructTest stTest)
        {
            try
            {
                var x = stTest.x;
                var y = stTest.y;

                return $"{x}_{y}";
            }
            catch(Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex.Message);
                return "";
            }
        }

        public string SimpleFunc()
        {
            return "Hello";
        }
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class UserInfos
    {
        public UserInfo[] UserInfoList;

        public override string ToString()
        {
            var msg = "";

            for (int i = 0; i < UserInfoList.Length; i++)
            {
                msg += UserInfoList[i].ToString() + Environment.NewLine;
            }
            return msg;
        }
    }

    [ComVisible(true)]
    [ClassInterface (ClassInterfaceType.AutoDual)]
    [DataContract]
    public class UserInfo
    {
        [DataMember(Name ="id")]
        public string Id { get; set; }

        [DataMember(Name ="userName")]
        public string UserName { get; set; }

        [DataMember(Name ="age")]
        public int Age { get; set; }

        [DataMember(Name ="enable")]
        public bool Enable { get; set; }

        [DataMember(Name ="food")]
        public string[] Food { get; set; }

        [DataMember(Name ="other")]
        public UserInfoOther Other { get; set; }

        public override string ToString()
        {
            var msg = "";

            msg += $"Id={Id}, Age={Age}, Enable={Enable}, " + Environment.NewLine;

            for (int i = 0; i < Food.Length; i++)
            {
                msg += $"Food={Food[i]}" + Environment.NewLine;
            }

            msg += $"Other={Other}" + Environment.NewLine;

            return msg;
        }
    }

    [DataContract]
    public class UserInfoOther
    {
        [DataMember(Name ="reserve")]
        public string Reserve { get; set; }

        public override string ToString()
        {
            return $"Reserve={Reserve}" + Environment.NewLine;
        }
    }
}
