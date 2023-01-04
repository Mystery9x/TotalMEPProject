using Autodesk.Revit.UI;
using TotalMEPProject.Commands.FireFighting;
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
                        CmdHolyUpdown.Run_UpdownElbowControl1();
                    }
                    break;

                case RequestId.HolyUpDown_DownElbowControl:
                    {
                        CmdHolyUpdown.Run_UpdownElbowControl1(true);
                    }
                    break;

                case RequestId.TwoLevelSmart_OK:
                    {
                        Cmd2LevelSmart.Process();
                    }
                    break;

                case RequestId.SprinklerUp_Aplly:
                    {
                        CmdSprinklerUpright.Process();
                    }
                    break;

                case RequestId.SprinklerDownType1_RUN:
                    {
                        CmdSprinklerDownright.ProcessType1();
                    }
                    break;

                case RequestId.SprinklerDownType2_RUN:
                    {
                        CmdSprinklerDownright.ProcessType2();
                    }
                    break;

                case RequestId.SprinklerDownType3_RUN:
                    {
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