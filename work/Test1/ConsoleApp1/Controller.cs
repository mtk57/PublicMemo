using CommonLib;

namespace ConsoleApp1
{
    public static class Controller
    {
        public static void Run(int param)
        {
            IService service = Factory.GetService(param);

            service.Validate(param);

            service.Read();
        }
    }
}
