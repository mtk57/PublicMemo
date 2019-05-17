using ConsoleApp1.DataAccess;
using ConsoleApp1.Factory;

namespace ConsoleApp1
{
    class Program
    {
        static int Main(string[] args)
        {
            var tool = PresentationFactory.GetPresentation(PresentationFactory.PresentaionType.Console);

            var result = tool.Run(new Param(args));

            return 0;
        }
    }
}
