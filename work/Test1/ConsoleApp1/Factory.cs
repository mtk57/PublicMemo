using CommonLib;
using System;
using System.Configuration;
using System.Reflection;

namespace ConsoleApp1
{
    public static class Factory
    {
        private const string KEY_FILE_NAME = @"FileName{0}";
        private const string KEY_CLASS_NAME = @"ClassName{0}";

        public static IService GetService(int param)
        {
            var fileNameKey = string.Format(KEY_FILE_NAME, param);
            var classNameKey = string.Format(KEY_CLASS_NAME, param);

            var valFileName = ConfigurationManager.AppSettings[fileNameKey];
            var valClassName = ConfigurationManager.AppSettings[classNameKey];

            Assembly asm = Assembly.LoadFrom(valFileName);
            Module mod = asm.GetModule(valFileName);
            System.Type type = mod.GetType(valClassName);
            if (type != null)
            {
                return (IService)Activator.CreateInstance(type);
            }
            return null;
        }
    }
}
