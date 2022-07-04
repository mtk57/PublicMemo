using System;
using System.Runtime.InteropServices;

namespace MyComLib
{
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("295D7DB1-FF92-4B2E-BD84-0B3802218F3D")]
    public interface IMyComLibArg
    {
        int myInt { get; set; }
        string myStr { get; set; }
        bool myBool { get; set; }
        bool SetStrArray([MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)] ref string[] arg);
    }

    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("117B82E7-6349-44FC-85D9-19239DCD7C62")]
    public interface IMyComLibMain
    {
        string Func(object obj);
        void RaiseError();
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [Guid("E6EA6F01-F2B4-41CC-8A86-B8AC58B6DF5B")]
    public class MyComLibArg : IMyComLibArg
    {
        public int myInt { get; set; }
        public string myStr { get; set; }
        public bool myBool { get; set; }

        public string[] myStrArray { get; private set; }

        public bool SetStrArray(ref string[] arg)
        {
            try
            {
                if (arg == null)
                {
                    return false;
                }

                myStrArray = arg;

                return true;

            }
            catch
            {
                return false;
            }
        }
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [Guid("FBF0B3A6-2CE4-404C-8325-9FEBCD01DF1C")]
    public class MyComLibMain : IMyComLibMain
    {
        public string Func(object arg)
        {
            try
            {
                if (arg == null)
                {
                    return "arg is null.";
                }

                if (!(arg is MyComLibArg))
                {
                    return "The argument types are different.";
                }

                var obj = arg as MyComLibArg;

                var primitiveArg = $"obj.myInt={obj.myInt}\nmyStr={obj.myStr}\nmyBool={obj.myBool}";

                var ary = obj.myStrArray;
                var ret_ary = "myStrArray=[";

                foreach (var s in ary)
                {
                    ret_ary += s + ", ";
                }
                ret_ary += "]"; 

                return $"success!\n{primitiveArg}\n{ret_ary}";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        public void RaiseError()
        {
            throw new Exception("Test Error");
        }
    }
}