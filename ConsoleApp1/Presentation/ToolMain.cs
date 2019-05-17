using ConsoleApp1.Factory;
using ConsoleApp1.Interface.DataAccess;

namespace ConsoleApp1.Presentation
{
    public class ToolMain : AbstractTool
    {
        protected override void PreProcess()
        {
        }

        protected override IResult MainProcess(IParam param)
        {
            var contoller = ConrollerFactory.GetContoller();

            return contoller.Run(param);
        }

        protected override void PostProcess()
        {
        }
    }
}
