using System.Runtime.InteropServices;

namespace MyComLib
{
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IMyComLibMain
    {
        string SimpleFunc();
    }

    [ClassInterface(ClassInterfaceType.None)]
    public class MyComLibMain : IMyComLibMain
    {
        public string SimpleFunc()
        {
            return "Hello 64bit Excel";
        }
    }
}