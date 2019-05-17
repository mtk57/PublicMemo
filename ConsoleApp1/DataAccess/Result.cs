using ConsoleApp1.Interface.DataAccess;

namespace ConsoleApp1.DataAccess
{
    public class Result : IResult
    {
        public int ResultCode { get; private set; }

        public IParam Param { get; private set; }

        private Result()
        {
        }

        public Result(int resultCode, IParam param)
        {
            this.ResultCode = resultCode;
            this.Param = param;
        }
    }
}
