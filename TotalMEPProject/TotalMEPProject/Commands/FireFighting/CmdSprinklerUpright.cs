using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TotalMEPProject.Ultis;

namespace TotalMEPProject.Commands.FireFighting
{
    [Transaction(TransactionMode.Manual)]
    public class CmdSprinklerUpright : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //if (App.CROK() != true)
            //    return Result.Failed;

            Global.UIApp = commandData.Application;
            Global.RVTApp = commandData.Application.Application;
            Global.UIDoc = commandData.Application.ActiveUIDocument;
            Global.AppCreation = commandData.Application.Application.Create;

            //Show form
            if (App.ShowSprinklerUpForm() == false)
                return Result.Cancelled;

            return Result.Succeeded;
        }

        public static Result Process()
        {
            List<FamilyInstance> sprinklers = sr.SelectSprinklers();
            if (sprinklers == null || sprinklers.Count == 0)
                return Result.Cancelled;

            List<Pipe> pipes = PickPipe();
            if (pipes == null || pipes.Count == 0)
                return Result.Cancelled;

            var pipeIds = (from Pipe p in pipes
                           where p.Id != ElementId.InvalidElementId
                           select p.Id).ToList();

            Transaction tran = new Transaction(Global.UIDoc.Document, "CreateConnector");
            tran.Start();

            //Find pipe
            foreach (FamilyInstance instance in sprinklers)
            {
                sr.cc(instance, pipeIds, true);
            }

            tran.Commit();
            return Result.Succeeded;
        }

        public static List<Pipe> PickPipe()
        {
            //Pick pipe
            List<Pipe> pipes = new List<Pipe>();
            try
            {
                var pickedObjs = Global.UIDoc.Selection.PickObjects(ObjectType.Element, new MEPCurveFilter/*PipeFilter*/(), "Pick pipes: ");

                foreach (Reference pickedObj in pickedObjs)
                {
                    var pipe = Global.UIDoc.Document.GetElement(pickedObj) as Pipe;

                    if (pipe != null)
                        pipes.Add(pipe);
                }
            }
            catch (System.Exception ex)
            {
            }

            return pipes;
        }
    }

    public class sr
    {
        public static List<FamilyInstance> SelectSprinklers()
        {
            List<FamilyInstance> list = new List<FamilyInstance>();
            try
            {
                var pickedObjs = Global.UIDoc.Selection.PickObjects(ObjectType.Element, new SprinklerFilter(), "Pick Sprinklers: ");

                foreach (Reference reference in pickedObjs)
                {
                    var familyInstance = Global.UIDoc.Document.GetElement(reference) as FamilyInstance;
                    if (familyInstance == null)
                        continue;
                    list.Add(familyInstance);
                }
            }
            catch (System.Exception ex)
            {
            }
            return list;
        }

        public static void XuLyDauong(Pipe pipe, out Pipe pipe2, XYZ pOn)
        {
            var curve = (pipe.Location as LocationCurve).Curve;

            //Create plane
            var p0 = curve.GetEndPoint(0);
            var p1 = curve.GetEndPoint(1);

            //Check co phai dau cuu hoa o gan dau cua ong ko : check trong pham vi 1m
            double kc_mm = 400 /*1000*/;
            double km_ft = Common.mmToFT * kc_mm;

            var p02d = new XYZ(p0.X, p0.Y, 0);
            var p12d = new XYZ(p1.X, p1.Y, 0);
            var pOn2d = new XYZ(pOn.X, pOn.Y, 0);

            var d1 = p02d.DistanceTo(pOn2d);
            var d2 = p12d.DistanceTo(pOn2d);

            bool isDauOng = false;
            int far = -1;
            if (d1 < km_ft)
            {
                //Check connect o dau ong
                //var c = Utils.GetConnectorClosestTo(pipe, p0);
                //if (c.IsConnected == false)
                if (ci(p0) == false)
                {
                    isDauOng = true;
                    far = 1;
                }
            }
            else if (d2 < km_ft)
            {
                //var c = Utils.GetConnectorClosestTo(pipe, p1);
                //if (c.IsConnected == false)
                if (ci(p1) == false)
                {
                    isDauOng = true;
                    far = 0;
                }
            }

            Pipe pipe1 = pipe;
            pipe2 = null;

            if (isDauOng == false)
            {
                s(pipe, pOn, out pipe1, out pipe2);
            }
            else if (far != -1)
            {
                if (far == 1)
                    (pipe1.Location as LocationCurve).Curve = Line.CreateBound(pOn, p1);
                else
                    (pipe1.Location as LocationCurve).Curve = Line.CreateBound(p0, pOn);
            }
        }

        public static void s(Pipe pipeOrigin, XYZ splitPoint, out Pipe pipe1, out Pipe pipe2)
        {
            var curve = (pipeOrigin.Location as LocationCurve).Curve;

            var p0 = curve.GetEndPoint(0);
            var p1 = curve.GetEndPoint(1);

            pipe1 = pipeOrigin;

            //Split
            (pipe1.Location as LocationCurve).Curve = Line.CreateBound(p0, splitPoint);

            var newPlace = new XYZ(0, 0, 0);
            var elemIds = ElementTransformUtils.CopyElement(
              Global.UIDoc.Document, pipeOrigin.Id, newPlace);

            pipe2 = Global.UIDoc.Document.GetElement(elemIds.ToList()[0]) as Pipe;

            (pipe2.Location as LocationCurve).Curve = Line.CreateBound(splitPoint, p1);

            //Find
            double ft = 0.001;
            var solid = Common.CreateCylindricalVolume(p1, ft, ft, false);
            if (solid != null)
            {
                //Find intersection wit fitting
                var fittingBuilt = new ElementId(BuiltInCategory.OST_PipeFitting);
                FilteredElementCollector collector = new FilteredElementCollector(Global.UIDoc.Document);
                collector.OfClass(typeof(FamilyInstance));
                collector.OfCategoryId(fittingBuilt);
                collector.WherePasses(new ElementIntersectsSolidFilter(solid)); // Apply intersection filter to find matches

                if (collector.GetElementCount() != 0)
                {
                    var elements = collector.ToElements();

                    var c1 = Common.GetConnectorClosestTo(pipe2, p1);

                    foreach (FamilyInstance fitting in elements)
                    {
                        var c11 = Common.GetConnectorClosestTo(fitting, p1);

                        if (c1 != null && c11 != null)
                        {
                            if (c1.Origin.DistanceTo(c11.Origin) < ft)
                            {
                                if (c1.IsConnectedTo(c11) == false)
                                    c1.ConnectTo(c11);
                            }
                        }
                    }
                }
            }
        }

        public static bool ci(XYZ point)
        {
            double ft = 0.001;
            var solid = Common.CreateCylindricalVolume(point, ft, ft, false);
            if (solid != null)
            {
                //Find intersection wit fitting
                var fittingBuilt = new ElementId(BuiltInCategory.OST_PipeFitting);
                FilteredElementCollector collector = new FilteredElementCollector(Global.UIDoc.Document);
                collector.OfClass(typeof(FamilyInstance));
                collector.OfCategoryId(fittingBuilt);
                collector.WherePasses(new ElementIntersectsSolidFilter(solid)); // Apply intersection filter to find matches

                if (collector.GetElementCount() != 0)
                {
                    return true;
                }
            }

            return false;
        }

        public static void ge(Pipe pipe, out Connector c0, out Connector c1)
        {
            var curve = (pipe.Location as LocationCurve).Curve;

            var p0 = curve.GetEndPoint(0);
            var p1 = curve.GetEndPoint(1);

            c0 = Common.GetConnectorClosestTo(pipe, p0);
            c1 = Common.GetConnectorClosestTo(pipe, p1);
        }

        public static bool cc(FamilyInstance instance, List<ElementId> selectedIds, bool isUp)
        {
            bool result = false;

            XYZ direction = isUp ? -XYZ.BasisZ : XYZ.BasisZ;

            //Set d = 25
            double d25 = 25;
            var dFt = Common.mmToFT * d25;

            var sprinkle_point = (instance.Location as LocationPoint).Point;

            XYZ newPlace = new XYZ(0, 0, 0);
            ICollection<ElementId> elemIds = null;

            var pipes = Common.FindElementsByDirection(Global.UIDoc.Document, ElementId.InvalidElementId, ElementId.InvalidElementId,
                ElementId.InvalidElementId, sprinkle_point, direction, BuiltInCategory.OST_PipeCurves, false, typeof(Pipe));

            if (pipes != null)
            {
                pipes = pipes.Where(item => selectedIds.Contains(item.Id)).ToList();
            }

            if (pipes == null || pipes.Count == 0)
            {
                double radius = 400; //mm
                var ft = Common.mmToFT * radius;

                var solid = Common.CreateCylindricalVolume(sprinkle_point, ft * 5, ft, !isUp);
                if (solid == null)
                    return result;

                ////Test////////////////////////////////////////////////////////////////////////
                //foreach (Edge edge in solid.Edges)
                //{
                //    var curve2 = edge.AsCurve();

                //    for (int i = 0; i < curve2.Tessellate().Count; i++)
                //    {
                //        XYZ p02 = curve2.Tessellate()[i];
                //        XYZ p12 = curve2.Tessellate()[(i + 1) % curve2.Tessellate().Count];

                //        Common.CreateModelLine(p02, p12);
                //    }
                //}

                //Find intersection
                FilteredElementCollector collector = new FilteredElementCollector(Global.UIDoc.Document, selectedIds);
                collector.OfClass(typeof(Pipe));
                collector.WherePasses(new ElementIntersectsSolidFilter(solid)); // Apply intersection filter to find matches

                if (collector.GetElementCount() == 0)
                    return result;
                var pipeList = collector.ToElements();

                //Global.m_uiDoc.Selection.SetElementIds(collector.ToElementIds());

                foreach (Pipe pipe in pipeList)
                {
                    var curve = (pipe.Location as LocationCurve).Curve;

                    var p0 = curve.GetEndPoint(0);
                    var p1 = curve.GetEndPoint(1);

                    //Move sprinker về điểm gần nhất với pipe
                    var sprinker2d = Common.ToPoint2D(sprinkle_point);
                    var p02d = Common.ToPoint2D(p0);
                    var p12d = Common.ToPoint2D(p1);

                    //                     if (p02d.DistanceTo(sprinker2d) < p12d.DistanceTo(sprinker2d))
                    //                     {
                    //                         sprinkle_point = new XYZ(p02d.X, p02d.Y, sprinkle_point.Z);
                    //                     }
                    //                     else
                    //                     {
                    //                         sprinkle_point = new XYZ(p12d.X, p12d.Y, sprinkle_point.Z);
                    //                     }

                    //(instance.Location as LocationPoint).Point = sprinkle_point;
                    //Global.m_uiDoc.Document.Regenerate();
                    //////////////////////////////////////////////////////////////////////////

                    var line2d = Line.CreateBound(new XYZ(p0.X, p0.Y, 0), new XYZ(p1.X, p1.Y, 0));
                    var v = line2d.Direction.CrossProduct(XYZ.BasisZ).Normalize();

                    double dTemp = 1000;

                    var lineTemp = Line.CreateUnbound(new XYZ(sprinkle_point.X, sprinkle_point.Y, 0), v * dTemp);

                    var p11 = lineTemp.Evaluate(dTemp, false);

                    lineTemp = Line.CreateUnbound(new XYZ(sprinkle_point.X, sprinkle_point.Y, 0), -v * dTemp);

                    var p22 = lineTemp.Evaluate(dTemp, false);

                    lineTemp = Line.CreateBound(p11, p22);

                    //expand
                    var index = line2d.GetEndPoint(0).DistanceTo(sprinkle_point) < line2d.GetEndPoint(1).DistanceTo(sprinkle_point) ? 0 : 1;
                    var curveExpand = Line.CreateUnbound(line2d.GetEndPoint(index), line2d.Direction * 100);

                    IntersectionResultArray arr = new IntersectionResultArray();
                    var inter = curveExpand.Intersect(lineTemp, out arr);
                    if (inter != SetComparisonResult.Overlap)
                        continue;

                    var p2d = arr.get_Item(0).XYZPoint;

                    var p3d = new XYZ(p2d.X, p2d.Y, sprinkle_point.Z);

                    var line3d = Line.CreateBound(p3d, new XYZ(p3d.X, p3d.Y, p3d.Z + (isUp ? -1 : 1) * dTemp));

                    var curveExtend3d_temp = Line.CreateUnbound(curve.GetEndPoint(index), (curve as Line).Direction * 100);

                    arr = new IntersectionResultArray();
                    inter = curveExtend3d_temp.Intersect(line3d, out arr);
                    if (inter != SetComparisonResult.Overlap)
                        continue;

                    var pOn = arr.get_Item(0).XYZPoint;

                    //Create pipe Z
                    newPlace = new XYZ(0, 0, 0);
                    elemIds = ElementTransformUtils.CopyElement(
                      Global.UIDoc.Document, pipe.Id, newPlace);

                    var newPipeZ = Global.UIDoc.Document.GetElement(elemIds.ToList()[0]) as Pipe;

                    lineTemp = Line.CreateBound(pOn, sprinkle_point);
                    var pcenter = lineTemp.Evaluate((lineTemp.GetEndParameter(0) + lineTemp.GetEndParameter(1)) / 2, false);

                    (newPipeZ.Location as LocationCurve).Curve = Line.CreateBound(pOn, pcenter);

                    newPipeZ.LookupParameter("Diameter").Set(dFt);

                    Global.UIDoc.Document.Regenerate();

                    Connector c1 = Common.GetConnectorClosestTo(pipe, pOn);
                    Connector c3 = Common.GetConnectorClosestTo(newPipeZ, pOn);

                    try
                    {
                        Global.UIDoc.Document.Create.NewElbowFitting(c1, c3);

                        result = true;
                    }
                    catch (System.Exception ex)
                    {
                    }

                    try
                    {
                        Connector c4 = Common.GetConnectorClosestTo(instance, pcenter);
                        Connector c5 = Common.GetConnectorClosestTo(newPipeZ, sprinkle_point);

                        Global.UIDoc.Document.Create.NewTransitionFitting(c5, c4);

                        result = true;
                    }
                    catch (System.Exception ex)
                    {
                    }
                }
            }
            else
            {
                foreach (Pipe pipe in pipes)
                {
                    var curve = (pipe.Location as LocationCurve).Curve;

                    //Create plane
                    var p0 = curve.GetEndPoint(0);
                    var p1 = curve.GetEndPoint(1);

                    //var plane = Plane.CreateByThreePoints(p0, p1, new XYZ(p1.X, p1.Y, p1.Z + 1));

                    //var sprinkle_point_1 = plane.ProjectOnto(sprinkle_point);
                    //if (sprinkle_point_1 == null)
                    //{
                    //    sprinkle_point_1 = sprinkle_point;//Chung to sprinkle_point nam giua pipe
                    //}

                    var proj = curve.Project(sprinkle_point);
                    var sprinkle_point_1 = proj.XYZPoint;

                    var lineZ = Line.CreateBound(sprinkle_point_1, new XYZ(sprinkle_point_1.X, sprinkle_point_1.Y, sprinkle_point_1.Z + 1000 * (isUp ? -1 : 1)));

                    XYZ pOn = null;
                    IntersectionResultArray array = new IntersectionResultArray();
                    var inter = lineZ.Intersect(curve, out array);
                    if (inter != SetComparisonResult.Overlap)
                    {
                        continue;
                    }
                    else
                        pOn = array.get_Item(0).XYZPoint;

                    //Move sprinkler ////////////////////////////////////////////////////////////////////////
                    sprinkle_point = new XYZ(pOn.X, pOn.Y, sprinkle_point.Z);
                    (instance.Location as LocationPoint).Point = sprinkle_point;
                    //////////////////////////////////////////////////////////////////////////

                    //Calculate point on curve

                    Pipe pipe1 = pipe;
                    Pipe pipe2 = null;
                    XuLyDauong(pipe, out pipe2, pOn);

                    //Create pipe Z
                    newPlace = new XYZ(0, 0, 0);
                    elemIds = ElementTransformUtils.CopyElement(
                      Global.UIDoc.Document, pipe.Id, newPlace);

                    var newPipeZ = Global.UIDoc.Document.GetElement(elemIds.ToList()[0]) as Pipe;

                    var lineTemp = Line.CreateBound(pOn, sprinkle_point);
                    var pcenter = lineTemp.Evaluate((lineTemp.GetEndParameter(0) + lineTemp.GetEndParameter(1)) / 2, false);

                    (newPipeZ.Location as LocationCurve).Curve = Line.CreateBound(pOn, pcenter);

                    newPipeZ.LookupParameter("Diameter").Set(dFt);

                    Global.UIDoc.Document.Regenerate();

                    Connector c1 = Common.GetConnectorClosestTo(pipe1, pOn);
                    Connector c3 = Common.GetConnectorClosestTo(newPipeZ, pOn);

                    if (pipe2 != null)
                    {
                        Connector c2 = Common.GetConnectorClosestTo(pipe2, pOn);

                        try
                        {
                            Global.UIDoc.Document.Create.NewTeeFitting(c1, c2, c3);

                            result = true;
                        }
                        catch (System.Exception ex)
                        {
                        }

                        //Add them Pipe sau khi cat
                        selectedIds.Add(pipe2.Id);
                    }
                    else
                    {
                        try
                        {
                            Global.UIDoc.Document.Create.NewElbowFitting(c1, c3);

                            result = true;
                        }
                        catch (System.Exception ex)
                        {
                        }
                    }

                    try
                    {
                        Connector c4 = Common.GetConnectorClosestTo(instance, pcenter);
                        Connector c5 = Common.GetConnectorClosestTo(newPipeZ, sprinkle_point);

                        Global.UIDoc.Document.Create.NewTransitionFitting(c5, c4);

                        result = true;
                    }
                    catch (System.Exception ex)
                    {
                    }

                    break;
                }
            }

            return result;
        }
    }

    public class SprinklerFilter : ISelectionFilter
    {
        public bool AllowElement(Element element)
        {
            return element != null && element.Category != null && (BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_Sprinklers;
        }

        public bool AllowReference(Reference refer, XYZ point)
        {
            return false;
        }
    }
}