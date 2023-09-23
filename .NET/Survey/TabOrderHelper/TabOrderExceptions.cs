namespace TabOrderHelper
{
    internal class ControlNotFoundException : System.Exception
    {
        private ControlNotFoundException()
        {
            // do nothing
        }

        public ControlNotFoundException(string message) : base(message)
        {
            // do nothing
        }
    }
}
