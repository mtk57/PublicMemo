using ConsoleApp1.Interface.Presentation;
using ConsoleApp1.Presentation;
using ConsoleApp1.Service.Exception;

namespace ConsoleApp1.Factory
{
    public static class PresentationFactory
    {
        public enum PresentaionType
        {
            Console,
            Window
        }

        public static IPresentation GetPresentation(PresentaionType type)
        {
            switch(type)
            {
                case PresentaionType.Console: return new ToolMain();
                default: throw new ToolException(-1, null, "Not support.");
            }
        }
    }
}
