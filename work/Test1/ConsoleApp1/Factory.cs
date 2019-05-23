using CommonLib;
using System;
using System.Configuration;
using System.Reflection;

namespace ConsoleApp1
{
    public static class Factory
    {
        public static IService GetService(int param)
        {
            if (param == 1)
            {
                string val1 = ConfigurationManager.AppSettings["key1"];

                Assembly asm = Assembly.LoadFrom(val1);
                Module mod = asm.GetModule("ServiceA.dll");
                System.Type type = mod.GetType("ServiceA.ServiceA");
                if (type != null)
                {
                    return (IService)Activator.CreateInstance(type);
                }
            }
            return null;
        }
    }
}
