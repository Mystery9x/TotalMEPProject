using Autodesk.Revit.UI;

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