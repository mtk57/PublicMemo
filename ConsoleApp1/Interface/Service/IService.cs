using ConsoleApp1.Interface.DataAccess;

namespace ConsoleApp1.Interface.Service
{
    public interface IService
    {
        IResult Execute(IParam param);
    }
}
