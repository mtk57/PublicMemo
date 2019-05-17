using ConsoleApp1.Interface.DataAccess;

namespace ConsoleApp1.Interface.Contoller
{
    public interface IContoller
    {
        IResult Run(IParam param);
    }
}
