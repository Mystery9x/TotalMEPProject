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

                case RequestId.HolyUpDown_PickObjects:
                    {
                        CmdHolyUpdown.Run_PickObjects();
                    }
                    break;

                case RequestId.HolyUpDown_OK:
                    {
                        CmdHolyUpdown.Run_OKRefreshDara();
                    }
                    break;

                case RequestId.HolyUpDown_UpStep:
                    {
                        CmdHolyUpdown.Run_UpDownStep();
                    }
                    break;

                case RequestId.HolyUpDown_DownStep:
                    {
                        CmdHolyUpdown.Run_UpDownStep(true);
                    }
                    break;

                case RequestId.HolyUpDown_UpElbowControl:
                    {
                        CmdHolyUpdown.Run_UpDownElbowControl();
                    }
                    break;

                case RequestId.HolyUpDown_DownElbowControl:
                    {
                        CmdHolyUpdown.Run_UpDownElbowControl(true);
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