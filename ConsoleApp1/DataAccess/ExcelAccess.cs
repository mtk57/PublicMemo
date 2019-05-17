using ConsoleApp1.Interface.DataAccess;

namespace ConsoleApp1.DataAccess
{
    public class ExcelAccess : IDataAccess
    {
        public object Value { get; private set; }

        private ExcelAccess()
        {
        }

        public ExcelAccess(IParam param)
        {
        }
    }
}
