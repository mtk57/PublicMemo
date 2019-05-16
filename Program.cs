using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            var datas = new DataContainer<Data>(new Data(args));

            var r = datas.Container.First().Value;
        }
    }

    public interface IDataContainer<T> where T : IData
    {
        IEnumerable<T> Container { get; }
        void Add<T>(T data);
        void Clear();
    }

    public interface IData
    {
        object Value { get; }
    }

    public class Data : IData
    {
        public object Value { get; private set; }

        public Data(object value)
        {
            this.Value = value;
        }
    }

    public class DataContainer<T> : IDataContainer<T> where T : IData
    {
        public IEnumerable<T> Container { get; private set; }

        public DataContainer(T data)
        {
            this.Container = new List<T>() { data };
        }

        public void Add<T>(T data)
        {
            (this.Container as IList<T>).Add(data);
        }

        public void Clear()
        {
            (this.Container as IList<T>).Clear();
        }
    }
}
