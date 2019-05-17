using ConsoleApp1.Factory;
using ConsoleApp1.Interface.Contoller;
using ConsoleApp1.Interface.DataAccess;

namespace ConsoleApp1.Controller
{
    public class ServiceContoller : IContoller
    {
        public IResult Run(IParam param)
        {
            var validataion = ServiceFactory.GetService(ServiceFactory.ServiceName.Validation);

            var converet = ServiceFactory.GetService(ServiceFactory.ServiceName.Convert);

            var output = ServiceFactory.GetService(ServiceFactory.ServiceName.Output);

            var result = validataion.Execute(param);

            result = converet.Execute(result.Param);

            result = output.Execute(result.Param);

            return result;
        }
    }
}
