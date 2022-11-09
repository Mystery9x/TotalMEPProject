using Autodesk.Revit.UI;
using TotalMEPProject.Commands.TotalMEP;

namespace TotalMEPProject.Request
{
    public class RequestHandler : IExternalEventHandler
    {
        public Request Request { get; } = new Request();

        /// <summary>
        /// Execute run
        /// </summary>
        public void Execute(UIApplication uiApp)
        {
            UIDocument uiDoc = uiApp.ActiveUIDocument;

            switch (Request.Take())
            {
                case RequestId.None:
                    {
                        return;
                    }

                case RequestId.VerticalMEP:
                    {
                        CmdVerticalMEP.o();
                    }
                    break;

                case RequestId.FastVertical:
                    {
                        CmdFastVertical.Process();
                    }
                    break;

                default:
                    {
                        break;
                    }
            }
            return;
        }

        public string GetName()
        {
            return "TotalMEP";
        }
    }
}