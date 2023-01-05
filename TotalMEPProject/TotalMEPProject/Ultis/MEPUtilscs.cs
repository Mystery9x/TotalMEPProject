using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using System;
using System.Linq;

namespace TotalMEPProject.Ultis
{
    public class MEPUtilscs
    {
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