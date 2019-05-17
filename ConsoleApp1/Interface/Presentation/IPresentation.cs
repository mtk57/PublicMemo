using ConsoleApp1.Interface.DataAccess;

namespace ConsoleApp1.Interface.Presentation
{
    public interface IPresentation
    {
        IResult Run(IParam param);
    }
}
