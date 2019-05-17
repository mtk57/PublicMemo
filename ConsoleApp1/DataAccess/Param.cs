using ConsoleApp1.Interface.DataAccess;

namespace ConsoleApp1.DataAccess
{
    public class Param : IParam
    {
        public object Value { get; private set; }

        private Param()
        { 
        }

        public Param(object value)
        {
            this.Value = value;
        }
    }
}
