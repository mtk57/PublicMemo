using ConsoleApp1.DataAccess;
using ConsoleApp1.Interface.DataAccess;
using ConsoleApp1.Interface.Service;

namespace ConsoleApp1.Service.Output
{
    public class OutputService : IService
    {
        public IResult Execute(IParam param)
        {
            return new Result(0, new Param(null));
        }
    }
}
