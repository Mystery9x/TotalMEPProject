using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TotalMEPProject.UI;
using TotalMEPProject.Ultis;

namespace TotalMEPProject.Commands.TotalMEP
{
    [Transaction(TransactionMode.Manual)]
    public class CmdVerticalMEP : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Global.UIApp = commandData.Application;
            Global.RVTApp = commandData.Application.Application;
            Global.UIDoc = commandData.Application.ActiveUIDocument;
            Global.AppCreation = commandData.Application.Application.Create;

            if (App.ShowVerticalMEPForm() == false)
            {
                return Result.Cancelled;
            }

            return Result.Succeeded;
        }

        public static Result o()
        {
            var form = App.verticalMEPForm;

            Transaction t = new Transaction(Global.UIDoc.Document, "a");

            form.Hide();

            while (true)
            {
                try
                {
                    ObjectSnapTypes snapTypes = ObjectSnapTypes.Nearest |
                        ObjectSnapTypes.Midpoints |
                        ObjectSnapTypes.Endpoints |
                        ObjectSnapTypes.Intersections |
                        ObjectSnapTypes.Centers |
                        ObjectSnapTypes.Perpendicular |
                        ObjectSnapTypes.Points;

                    XYZ point = Global.UIDoc.Selection.PickPoint(snapTypes, "Select an point: ");

                    t.Start();

                    double startZ = 0;
                    double endZ = 0;

                    ce(form, out startZ, out endZ);

                    XYZ start = new XYZ(point.X, point.Y, startZ);

                    XYZ end = new XYZ(point.X, point.Y, endZ);

                    ElementId systemTypeId = ElementId.InvalidElementId;
                    if (form.MEPType_ == MEPType.Pipe)
                    {
                        systemTypeId = form.SystemType;
                    }
                    else if (form.MEPType_ == MEPType.Round_Duct)
                    {
                        systemTypeId = form.SystemType;
                    }
                    else if (form.MEPType_ == MEPType.Oval_Duct || form.MEPType_ == MEPType.Rectangular_Duct)
                    {
                        systemTypeId = form.SystemType;
                    }

                    var mepNew = err(start, end, form.FamilyType, systemTypeId, form.LevelBottomId);

                    if (mepNew != null)
                    {
                        //Set parameter
                        double height = 0; //inch
                        double width = 0;
                        if (form.MEPType_ == MEPType.Pipe || form.MEPType_ == MEPType.Round_Duct)
                        {
                            width = (form.MEPSize_ as MEPSize).NominalDiameter;
                            mepNew.LookupParameter("Diameter").Set(width);

                            height = (form.MEPSize_ as MEPSize).NominalDiameter;
                        }
                        else if (form.MEPType_ == MEPType.Conduit)
                        {
                            width = (form.MEPSize_ as ConduitSize).NominalDiameter;
                            mepNew.LookupParameter("Diameter(Trade Size)").Set(width);

                            height = (form.MEPSize_ as ConduitSize).NominalDiameter;
                        }
                        else
                        {
                            width = form.MEP_Width * Common.mmToFT;
                            mepNew.LookupParameter("Width").Set(width);
                            mepNew.LookupParameter("Height").Set(form.MEP_Height * Common.mmToFT);

                            height = form.MEP_Height * Common.mmToFT;
                        }

                        if (form.MEPType_ == MEPType.CableTray || form.MEPType_ == MEPType.Conduit && form.ServiceType != string.Empty)
                        {
                            mepNew.LookupParameter("Service Type").Set(form.ServiceType);
                        }
                    }
                    t.Commit();
                }
                catch (System.Exception ex)
                {
                    string mess = ex.Message;

                    if (mess.Contains("The user aborted the pick operation."))
                    {
                        form.TopMost = true;
                        form.Show();
                    }
                    break;
                }
            }

            if (t.HasStarted())
            {
                t.RollBack();
            }

            return Result.Succeeded;
        }

        private static void ce(VerticalMEPForm form, out double startZ, out double endZ)
        {
            var top = Global.UIDoc.Document.GetElement(form.LevelTopId) as Level;
            var bottom = Global.UIDoc.Document.GetElement(form.LevelBottomId) as Level;

            startZ = bottom.Elevation + form.OffsetBottom * Common.mmToFT;
            endZ = top.Elevation + form.OffsetTop * Common.mmToFT;
        }

        public static MEPCurve err(XYZ start, XYZ end, ElementId elementTypeId, ElementId systemTypeId, ElementId levelId)
        {
            ElementType elementType = Global.UIDoc.Document.GetElement(elementTypeId) as ElementType;

            MEPCurve mepCurve = null;

            if (elementType is Autodesk.Revit.DB.Mechanical.DuctType)
            {
                mepCurve = Autodesk.Revit.DB.Mechanical.Duct.Create(Global.UIDoc.Document, systemTypeId, elementTypeId, levelId, start, end);
            }
            else if (elementType is PipeType)
            {
                mepCurve = Pipe.Create(Global.UIDoc.Document, systemTypeId, elementTypeId, levelId, start, end);
            }
            else if (elementType is CableTrayType)
            {
                mepCurve = CableTray.Create(Global.UIDoc.Document, elementTypeId, start, end, levelId);
            }
            else if (elementType is ConduitType)
            {
                mepCurve = Conduit.Create(Global.UIDoc.Document, elementTypeId, start, end, levelId);
            }
            return mepCurve;
        }
    }
}