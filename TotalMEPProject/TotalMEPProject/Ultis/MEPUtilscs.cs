using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.UI;
using System;
using System.Linq;

namespace TotalMEPProject.Ultis
{
    public class MEPUtilscs
    {
        public static Pipe PickPipe(UIDocument uidoc, ISelectionFilter selectionFilter, string statusPrompt)
        {
            Pipe retval = null;
            try
            {
                if (uidoc == null || uidoc.Document == null)
                    return retval;

                Reference refEle = null;

                if (selectionFilter == null)
                    refEle = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, statusPrompt);
                else
                    refEle = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, selectionFilter, statusPrompt);

                if (refEle != null)
                {
                    Element ele = uidoc.Document.GetElement(refEle);
                    if (ele != null && ele is Pipe pipe)
                        return pipe;
                }
                return retval;
            }
            catch (Exception)
            {
                return retval;
            }
        }

        public static Element CE(MEPCurve mep1, MEPCurve mep2, XYZ point)
        {
            var c3 = Common.GetConnectorClosestTo(mep1, point);
            var c4 = Common.GetConnectorClosestTo(mep2, point);

            try
            {
                var elbow = Global.UIDoc.Document.Create.NewElbowFitting(c3, c4);

                return elbow;
            }
            catch (System.Exception ex)
            {
                return null;
            }
        }

        public static MEPCurve CC(MEPCurve pipeTempMain1, Line line)
        {
            var newPlace = new XYZ(0, 0, 0);
            var elemIds = ElementTransformUtils.CopyElement(
               Global.UIDoc.Document, pipeTempMain1.Id, newPlace);

            var curveMain = (pipeTempMain1.Location as LocationCurve).Curve as Line;

            var main2 = Global.UIDoc.Document.GetElement(elemIds.ToList()[0]) as MEPCurve;

            (main2.Location as LocationCurve).Curve = line;

            if (Common.IsVertical(line.Direction))
            {
                if (pipeTempMain1 as Duct != null || pipeTempMain1 as CableTray != null)
                {
                    var angle = XYZ.BasisX.AngleOnPlaneTo(curveMain.Direction, XYZ.BasisZ);

                    main2.Location.Rotate(Line.CreateUnbound(line.GetEndPoint(0), XYZ.BasisZ), angle + Math.PI / 2);
                }
            }

            return main2;
        }

        public static MEPCurve CC1(MEPCurve pipeTempMain1, Line line)
        {
            var newPlace = new XYZ(0, 0, 0);
            var elemIds = ElementTransformUtils.CopyElement(
               Global.UIDoc.Document, pipeTempMain1.Id, newPlace);

            var curveMain = (pipeTempMain1.Location as LocationCurve).Curve as Line;

            var main2 = Global.UIDoc.Document.GetElement(elemIds.ToList()[0]) as MEPCurve;

            (main2.Location as LocationCurve).Curve = line;

            if (Common.IsVertical(line.Direction))
            {
                if (pipeTempMain1 as Duct != null || pipeTempMain1 as CableTray != null)
                {
                    var angle = XYZ.BasisX.AngleOnPlaneTo(curveMain.Direction, XYZ.BasisZ);

                    main2.Location.Rotate(Line.CreateUnbound(line.GetEndPoint(0), XYZ.BasisZ), angle + Math.PI / 2);
                }
            }

            return main2;
        }
    }
}