using ConsoleApp1.DataAccess;
using ConsoleApp1.Interface.DataAccess;
using ConsoleApp1.Interface.Exception;

namespace ConsoleApp1.Service.Exception
{
    public sealed class ToolException : System.Exception, IToolException
    {
        public IResult Result { get; private set; }

        private ToolException()
        {
        }

        public ToolException(int code, IParam param, string message) : base(message)
        {
            this.Result = new Result(code, param);
        }
    }
}
