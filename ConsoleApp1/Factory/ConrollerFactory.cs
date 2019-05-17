using ConsoleApp1.Controller;
using ConsoleApp1.Interface.Contoller;

namespace ConsoleApp1.Factory
{
    public static class ConrollerFactory
    {
        public static IContoller GetContoller()
        {
            return new ServiceContoller();
        }
    }
}
