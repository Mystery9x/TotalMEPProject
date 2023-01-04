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
using TotalMEPProject.UI.FireFightingUI;
using TotalMEPProject.Ultis;

namespace TotalMEPProject.Commands.FireFighting
{
    [Transaction(TransactionMode.Manual)]
    public class CmdSprinklerDownright : IExternalCommand
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
            if (App.ShowSprinklerDownForm() == false)
                return Result.Cancelled;

            return Result.Succeeded;
        }

        #region Type1

        public static Result ProcessType1()
        {
            List<FamilyInstance> sprinklers = sr.SelectSprinklers(App.m_SprinklerDownForm.isD15);
            if (sprinklers == null || sprinklers.Count == 0)
                return Result.Cancelled;

            List<Pipe> pipes = PickPipes();
            if (pipes == null || pipes.Count == 0)
                return Result.Cancelled;

            var pipeIds = (from Pipe p in pipes
                           where p.Id != ElementId.InvalidElementId
                           select p.Id).ToList();

            var height = App.m_SprinklerDownForm.Height_;

            Transaction tran = new Transaction(Global.UIDoc.Document, "CreateConnector");
            tran.Start();

            double radius = 500/*400*/; //mm : sua thanh 500 theo yeu cau cua a Cuong
            var ft = Common.mmToFT * radius;

            //Find pipe
            foreach (FamilyInstance instance in sprinklers)
            {
                var sprinkle_point = (instance.Location as LocationPoint).Point;

                //Check connect
                var connects = instance.MEPModel.ConnectorManager.Connectors;

                if (connects.Size == 0)
                    continue;

                var connect = instance.MEPModel.ConnectorManager.Lookup(1);

                if (connect.IsConnected == true)
                    continue;

                var solid = Common.CreateCylindricalVolume(sprinkle_point, ft * 5, ft, true);
                if (solid == null)
                    continue;

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

                ////////////////////////////////////////////////////////////////////////////

                //Find intersection

                FilteredElementCollector collector = new FilteredElementCollector(Global.UIDoc.Document, pipeIds);
                collector.OfClass(typeof(Pipe));
                collector.WherePasses(new ElementIntersectsSolidFilter(solid)); // Apply intersection filter to find matches
                if (collector.GetElementCount() == 0)
                    continue;
                var pipeList = collector.ToElements();

                //Global.m_uiDoc.Selection.SetElementIds(collector.ToElementIds());

                bool split = false;
                var pipe = pp(pipeList.ToList(), sprinkle_point, out split);

                var curve = (pipe.Location as LocationCurve).Curve;

                var p0 = curve.GetEndPoint(0);
                var p1 = curve.GetEndPoint(1);

                double dTemp = 1000;

                var line2d = Line.CreateBound(new XYZ(p0.X, p0.Y, 0), new XYZ(p1.X, p1.Y, 0));
                var v = line2d.Direction.CrossProduct(XYZ.BasisZ).Normalize();

                var lineTemp = Line.CreateUnbound(new XYZ(sprinkle_point.X, sprinkle_point.Y, 0), v * dTemp);

                var p11 = lineTemp.Evaluate(dTemp, false);

                lineTemp = Line.CreateUnbound(new XYZ(sprinkle_point.X, sprinkle_point.Y, 0), -v * dTemp);

                var p22 = lineTemp.Evaluate(dTemp, false);

                lineTemp = Line.CreateBound(p11, p22);

                //Common.CreateModelLine(line2d.GetEndPoint(0), line2d.GetEndPoint(1));
                //Common.CreateModelLine(p11, p22);

                XYZ newPlace = new XYZ(0, 0, 0);
                ICollection<ElementId> elemIds = null;
                var pipe1 = pipe;
                Pipe pipe2 = null;
                XYZ p = null;

                IntersectionResultArray arr = new IntersectionResultArray();
                if (split == false)
                {
                    //truong hop o dau ong

                    //expand
                    var index = line2d.GetEndPoint(0).DistanceTo(sprinkle_point) < line2d.GetEndPoint(1).DistanceTo(sprinkle_point) ? 0 : 1;
                    var curveExpand = Line.CreateUnbound(line2d.GetEndPoint(index), line2d.Direction * 100);

                    var inter = curveExpand.Intersect(lineTemp, out arr);
                    if (inter != SetComparisonResult.Overlap)
                        continue;

                    var p2d = arr.get_Item(0).XYZPoint;

                    var p3d = new XYZ(p2d.X, p2d.Y, sprinkle_point.Z);

                    var line3d = Line.CreateBound(p3d, new XYZ(p3d.X, p3d.Y, p3d.Z + dTemp));

                    var curveExtend3d_temp = Line.CreateUnbound(curve.GetEndPoint(index), (curve as Line).Direction * 100);

                    arr = new IntersectionResultArray();
                    inter = curveExtend3d_temp.Intersect(line3d, out arr);
                    if (inter != SetComparisonResult.Overlap)
                        continue;

                    p = arr.get_Item(0).XYZPoint;
                }
                else
                {
                    var inter = lineTemp.Intersect(line2d, out arr);
                    if (inter != SetComparisonResult.Overlap)
                        continue;

                    var p2d = arr.get_Item(0).XYZPoint;

                    var p3d = new XYZ(p2d.X, p2d.Y, sprinkle_point.Z);

                    var line3d = Line.CreateBound(p3d, new XYZ(p3d.X, p3d.Y, p3d.Z + dTemp));

                    arr = new IntersectionResultArray();
                    inter = curve.Intersect(line3d, out arr);
                    if (inter != SetComparisonResult.Overlap)
                        continue;

                    p = arr.get_Item(0).XYZPoint;

                    pipe2 = null;
                    sr.XuLyDauong(pipe, out pipe2, p);

                    if (pipe2 != null)
                    {
                        pipeIds.Add(pipe2.Id);
                    }
                }

                //Set d = 25
                //double d25 = 25;
                var dFt = Common.mmToFT * App.m_SprinklerDownForm.PipeSize;

                var ft_h = Common.mmToFT * height;

                newPlace = new XYZ(0, 0, 0);
                elemIds = ElementTransformUtils.CopyElement(
                 Global.UIDoc.Document, pipe1.Id, newPlace);

                var pipe_v1 = Global.UIDoc.Document.GetElement(elemIds.ToList()[0]) as Pipe;

                var line_v1 = Line.CreateUnbound(p, XYZ.BasisZ * ft_h * 2);

                line_v1 = Line.CreateBound(p, line_v1.Evaluate(ft_h, false));
                (pipe_v1.Location as LocationCurve).Curve = line_v1;

                pipe_v1.LookupParameter("Diameter").Set(dFt);

                //Connect
                try
                {
                    var c1 = Common.GetConnectorClosestTo(pipe1, p);
                    var c3 = Common.GetConnectorClosestTo(pipe_v1, p);

                    if (pipe2 != null)
                    {
                        var c2 = Common.GetConnectorClosestTo(pipe2, p);
                        var fitting = Global.UIDoc.Document.Create.NewTeeFitting(c1, c2, c3);
                    }
                    else
                    {
                        Global.UIDoc.Document.Create.NewElbowFitting(c1, c3);
                    }
                }
                catch (System.Exception ex)
                {
                }

                //Hor
                var v_v = (new XYZ(sprinkle_point.X, sprinkle_point.Y, 0) - new XYZ(p.X, p.Y, 0)).Normalize();

                var ft_v = (new XYZ(sprinkle_point.X, sprinkle_point.Y, 0) - new XYZ(p.X, p.Y, 0)).GetLength();

                var line_Extend = Line.CreateUnbound(line_v1.GetEndPoint(1), ft_v * v_v * 2);

                newPlace = new XYZ(0, 0, 0);
                elemIds = ElementTransformUtils.CopyElement(
                 Global.UIDoc.Document, pipe1.Id, newPlace);

                var pipe_hor = Global.UIDoc.Document.GetElement(elemIds.ToList()[0]) as Pipe;
                var line_hor = Line.CreateBound(line_v1.GetEndPoint(1), line_Extend.Evaluate(ft_v, false));

                (pipe_hor.Location as LocationCurve).Curve = line_hor;

                pipe_hor.LookupParameter("Diameter").Set(dFt);

                try
                {
                    var c1 = Common.GetConnectorClosestTo(pipe_v1, line_v1.GetEndPoint(1));
                    var c2 = Common.GetConnectorClosestTo(pipe_hor, line_v1.GetEndPoint(1));

                    Global.UIDoc.Document.Create.NewElbowFitting(c1, c2);
                }
                catch (System.Exception ex)
                {
                }

                //Vertical 2
                var line_v2 = Line.CreateBound(line_hor.GetEndPoint(1), sprinkle_point);

                newPlace = new XYZ(0, 0, 0);
                elemIds = ElementTransformUtils.CopyElement(
                 Global.UIDoc.Document, pipe1.Id, newPlace);

                var pipe_v2 = Global.UIDoc.Document.GetElement(elemIds.ToList()[0]) as Pipe;
                var center = line_v2.Evaluate((line_v2.GetEndParameter(0) + line_v2.GetEndParameter(1)) / 2, false);

                (pipe_v2.Location as LocationCurve).Curve = Line.CreateBound(line_v2.GetEndPoint(0), center);

                pipe_v2.LookupParameter("Diameter").Set(dFt);

                //Connect
                try
                {
                    var c1 = Common.GetConnectorClosestTo(pipe_hor, line_hor.GetEndPoint(1));
                    var c2 = Common.GetConnectorClosestTo(pipe_v2, line_hor.GetEndPoint(1));

                    Global.UIDoc.Document.Create.NewElbowFitting(c1, c2);
                }
                catch (System.Exception ex)
                {
                }

                try
                {
                    var c1 = Common.GetConnectorClosestTo(pipe_v2, center);
                    var c2 = Common.GetConnectorClosestTo(instance, center);

                    Global.UIDoc.Document.Create.NewTransitionFitting(c1, c2);
                }
                catch (System.Exception ex)
                {
                }
            }

            tran.Commit();

            return Result.Succeeded;
        }

        public static List<Pipe> PickPipes()
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

        private static Pipe pp(List<Element> pippes, XYZ sprinkle_point, out bool bSplit)
        {
            bSplit = false;

            if (pippes.Count == 0)
                return null;

            Pipe pipeNear = null;
            foreach (Element pipe in pippes)
            {
                if (pipe as Pipe == null)

                    continue;
                var curve = (pipe.Location as LocationCurve).Curve;

                if (curve is Line == false)
                    continue;

                var d = (curve as Line).Direction;

                if (Common.IsParallel(d, XYZ.BasisZ, 0))
                    continue;

                var project = curve.Project(sprinkle_point);
                if (project == null)
                    continue;

                var p = project.XYZPoint;
                if (p.DistanceTo(curve.GetEndPoint(0)) != 0 && p.DistanceTo(curve.GetEndPoint(1)) != 0)
                {
                    bSplit = true;
                    return pipe as Pipe;
                }
                else
                {
                    pipeNear = pipe as Pipe;
                }
            }
            return pipeNear;
        }

        #endregion Type1

        #region Type2

        public static Result ProcessType2()
        {
            List<FamilyInstance> sprinklers = sr.SelectSprinklers(App.m_SprinklerDownForm.isD15);
            if (sprinklers == null || sprinklers.Count == 0)
                return Result.Cancelled;

            List<Pipe> pipes = PickPipes();
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
                cc(instance, pipeIds, App.m_SprinklerDownForm.FamilyType, App.m_SprinklerDownForm.PipeSize, false);
            }

            tran.Commit();

            return Result.Succeeded;
        }

        public static bool cc(FamilyInstance instance, List<ElementId> selectedIds, ElementId pipeTypeId, double pipeSize, bool isUp)
        {
            bool result = false;

            XYZ direction = isUp ? -XYZ.BasisZ : XYZ.BasisZ;

            ////Set d = 25
            //double d25 = 25;
            var dFt = Common.mmToFT * pipeSize;

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

                    var resultPipe = curve.Project(sprinkle_point);

                    XYZ pointProject = resultPipe.XYZPoint;
                    XYZ pointProject2d = Common.ToPoint2D(pointProject);

                    var distancePoint2d = pointProject2d.DistanceTo(sprinker2d);
                    if (!Common.IsEqual(distancePoint2d, 0))
                    {
                        var vector = pointProject2d - sprinker2d;

                        ElementTransformUtils.MoveElement(Global.UIDoc.Document, instance.Id, vector.Normalize() * pointProject2d.DistanceTo(sprinker2d));

                        Global.UIDoc.Document.Regenerate();

                        sprinkle_point = (instance.Location as LocationPoint).Point;
                    }

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

                    var height = UnitUtils.ConvertToInternalUnits(App.m_SprinklerDownForm.Height_, DisplayUnitType.DUT_MILLIMETERS);

                    pcenter = new XYZ(pcenter.X, pcenter.Y, pOn.Z - height);

                    (newPipeZ.Location as LocationCurve).Curve = Line.CreateBound(pOn, pcenter);

                    newPipeZ.LookupParameter("Diameter").Set(dFt);

                    var pipeType = Global.UIDoc.Document.GetElement(pipeTypeId) as PipeType;
                    if (pipeType != null)
                        newPipeZ.PipeType = pipeType;

                    Global.UIDoc.Document.Regenerate();

                    Connector c1 = Common.GetConnectorClosestTo(pipe, pOn);
                    Connector c3 = Common.GetConnectorClosestTo(newPipeZ, pOn);

                    try
                    {
                        //Global.UIDoc.Document.Create.NewElbowFitting(c1, c3);

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
                    sr.XuLyDauong(pipe, out pipe2, pOn);

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

        #endregion Type2
    }
}