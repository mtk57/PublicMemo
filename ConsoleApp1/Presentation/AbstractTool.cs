using ConsoleApp1.Interface.DataAccess;
using ConsoleApp1.Interface.Presentation;

namespace ConsoleApp1.Presentation
{
    public abstract class AbstractTool : IPresentation
    {
        protected abstract void PreProcess();

        protected abstract IResult MainProcess(IParam param);

        protected abstract void PostProcess();

        public virtual IResult Run(IParam param)
        {
            PreProcess();

            var result = MainProcess(param);

            PostProcess();

            return result;
        }
    }
}
