using ConsoleApp1.Interface.DataAccess;

namespace ConsoleApp1.DataAccess
{
    public class ExcelAccess : IDataAccess
    {
        // TODO:
        // 本来であれば、生のDataSetは隠蔽して、DataTableをラップしたI/F(ITable)を返すべき。
        // ITableには便利メソッドを追加していく
        //   Ex.
        //   ・指定した範囲をITableで返す。(1行目を列名とするか否かを指定可)
        //   ・開始位置から空列,空行までの範囲をITableで返す。(1行目を列名とするか否かを指定可)
        //   ・指定した条件を満たす列名だけをフィルタリングする。(行は全てある)
        //   ・指定した1列の条件を満たす行だけをフィルタリングする。(列は全てある)
        public object Value { get; private set; }

        private ExcelAccess()
        {
        }

        public ExcelAccess(IParam param)
        {
        }
    }
}
