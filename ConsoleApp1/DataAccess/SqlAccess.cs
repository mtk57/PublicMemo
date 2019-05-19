using ConsoleApp1.Interface.DataAccess;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Linq;

namespace ConsoleApp1.DataAccess
{
    // TODO:
    // 本来であればSqlAccessはabstractにして、SubClassとしてPostgreSqlAccess, SQLServerSqlAccessというようにすべき
    // ->利用者にDBの種類は意識させない


    // [Codeイメージ]
    // var sql = new SqlAccess("SQLファイル出力パス");
    // 
    // var deleteList = new List<string>();
    // var insertList = new List<string>();
    //
    // var layoutIds = GetAllLayoutId();
    // 
    // // Delete SQL作成
    // deleteList.Add( sql.CreateDelete( SqlAccess.Table.A, SqlAccess.Table.Column.A1, layoutIds ));
    //
    // var insertA = GetInsertA(); // tableAのプロパティに値を設定
    // 
    // // Insert SQL作成
    // insertList.Add( sql.CreateInsert( insertA );
    //
    // // SQL出力
    // sql.Output(Encoding.UTF8);
    //
    public class SqlAccess : IDataAccess
    {
        public const string TEST_SQL = @"
INSERT INTO {0}
({1})
VALUES
({2});
";

        public object Value { get; private set; }

        private SqlAccess()
        {
        }

        public SqlAccess(IParam param)
        {
        }

        public static string CreateInsert(IDBTable table)
        {
            return string.Format(
                TEST_SQL,
                table.TABLE_NAME,
                GetFieldNames(table),
                GetFieldValues(table)
                );
        }

        private static string GetFieldNames(IDBTable table)
        {
            var pi = GetPropertyInfo(table);

            var fields = pi.Select(y => y.Name);

            return string.Join(",", fields);
        }

        private static string GetFieldValues(IDBTable table)
        {
            var pi = GetPropertyInfo(table);

            var values = pi.Select(x => 
            {
                var value = x.GetValue(table).ToString();
                return (x.PropertyType.Name == "String") ? string.Format($"'{value}'") : value;
            });

            return string.Join(",", values);
        }

        private static IEnumerable<PropertyInfo> GetPropertyInfo(object obj)
        {
            return obj.GetType().GetProperties().Where(x => x.Name != "TABLE_NAME");
        }
    }

    public interface IDBTable
    {
        string TABLE_NAME { get; }
    }

    public class TestTableA : IDBTable
    {
        public string TABLE_NAME { get; } = "TableA";

        public int FieldA { get; set; }
        public bool FieldB { get; set; }
        public string FieldC { get; set; }

        private TestTableA() { }

        public TestTableA(int fieldA, bool fieldB, string fieldC)
        {
            this.FieldA = fieldA;
            this.FieldB = fieldB;
            this.FieldC = fieldC;
        }
    }
}
