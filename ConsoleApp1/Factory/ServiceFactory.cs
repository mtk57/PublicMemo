using ConsoleApp1.Interface.Service;
using ConsoleApp1.Service.Convert;
using ConsoleApp1.Service.Exception;
using ConsoleApp1.Service.Output;
using ConsoleApp1.Service.Validation;

namespace ConsoleApp1.Factory
{
    public static class ServiceFactory
    {
        public enum ServiceName
        {
            Validation,
            Convert,
            Output
        }

        public static IService GetService(ServiceName serviceName)
        {
            switch (serviceName)
            {
                case ServiceName.Validation: return new ValidationService();
                case ServiceName.Convert: return new ConvertService();
                case ServiceName.Output: return new OutputService();
                default: throw new ToolException(-1, null, "Service not found");
            }
        }
    }
}
