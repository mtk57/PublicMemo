using System;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Controller.Run(1);
            }
            catch(Exception ex)
            {
                Console.WriteLine(string.Format("Msg={0}, Stack={1}", ex.Message, ex.StackTrace));
            }
        }
    }
}
