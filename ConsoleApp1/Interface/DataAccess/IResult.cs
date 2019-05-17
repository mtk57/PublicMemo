namespace ConsoleApp1.Interface.DataAccess
{
    public interface IResult
    {
        int ResultCode { get; }
        IParam Param { get; }
    }
}
