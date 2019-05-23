using CommonLib;

namespace ServiceA
{
    public class ServiceA : IService
    {
        public void Validate(object param)
        {
            System.Diagnostics.Debug.WriteLine("Validate");
        }

        public void Read()
        {
            System.Diagnostics.Debug.WriteLine("Validate");
        }
    }
}
