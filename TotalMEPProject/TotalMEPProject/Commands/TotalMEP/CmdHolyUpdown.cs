using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TotalMEPProject.UI;
using TotalMEPProject.Ultis;

namespace TotalMEPProject.Commands.TotalMEP
{
    [Transaction(TransactionMode.Manual)]
    public class CmdHolyUpdown : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Global.UIApp = commandData.Application;
            Global.RVTApp = commandData.Application.Application;
            Global.UIDoc = commandData.Application.ActiveUIDocument;
            Global.AppCreation = commandData.Application.Application.Create;

            if (App.ShowHolyUpDownForm() == false)
            {
                return Result.Cancelled;
            };
            return Result.Succeeded;
        }

        public static Result Run_PickObjects()
        {
            return Result.Succeeded;
        }
    }
}