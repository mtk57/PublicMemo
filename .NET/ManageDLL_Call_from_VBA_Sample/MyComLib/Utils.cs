using System.IO;
using System.Reflection;

namespace MyComLib.Common
{
    internal static class Utils
    {
        public static string GetMyDir()
        {
            return Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        }
    }
}
