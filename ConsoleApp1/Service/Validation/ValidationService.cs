using ConsoleApp1.DataAccess;
using ConsoleApp1.Interface.DataAccess;
using ConsoleApp1.Interface.Service;

namespace ConsoleApp1.Service.Validation
{
    public class ValidationService : IService
    {
        public IResult Execute(IParam param)
        {
            return new Result(0, new Param(null));
        }
    }
}
