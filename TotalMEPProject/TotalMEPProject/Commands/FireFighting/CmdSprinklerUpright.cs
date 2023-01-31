using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Interop;
using TotalMEPProject.Services;
using TotalMEPProject.UI.BeginUI;
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
            try
            {
                if (App.m_SprinkerUpForm != null && App.m_SprinkerUpForm.IsDisposed == false)
                {
                    App.m_SprinkerUpForm.Hide();
                }

                List<FamilyInstance> sprinklers = sr.SelectSprinklers(App.m_SprinkerUpForm.isD15);
                if (sprinklers == null || sprinklers.Count == 0)
                    return Result.Cancelled;

                List<Pipe> pipes = PickPipe();
                if (pipes == null || pipes.Count == 0)
                    return Result.Cancelled;

                var pipeIds = (from Pipe p in pipes
                               where p.Id != ElementId.InvalidElementId
                               select p.Id).ToList();

                // Status cancel export : default = false
                bool isCancelExport = false;

                // Count type imported
                int nCount = 0;

                System.Diagnostics.Process process = System.Diagnostics.Process.GetCurrentProcess();
                IntPtr intPtr = process.MainWindowHandle;
                ViewSingleProgressBar progressBar = new ViewSingleProgressBar("Sprinkler Up", "Process : ");
                progressBar.prgSingle.Minimum = 1;
                progressBar.prgSingle.Maximum = sprinklers.Count;
                progressBar.prgSingle.Value = 1;
                WindowInteropHelper helper = new WindowInteropHelper(progressBar);
                helper.Owner = intPtr;
                progressBar.Show();

                TransactionGroup tranGr = new TransactionGroup(Global.UIDoc.Document, "CreateConnector");
                tranGr.Start();

                //Find pipe
                try
                {
                    foreach (FamilyInstance instance in sprinklers)
                    {
                        double dPercent = 0.0;
                        try
                        {
                            Transaction tran = new Transaction(Global.UIDoc.Document, "CreateConnector");
                            tran.Start();
                            sr.cc(tran, instance, pipeIds, App.m_SprinkerUpForm.isConnectNipple, App.m_SprinkerUpForm.isConnectTee, App.m_SprinkerUpForm.fmlNipple, App.m_SprinkerUpForm.PipeSize, App.m_SprinkerUpForm.FamilyType, true);

                            tran.Commit();
                        }
                        catch (Exception)
                        { }

                        // If click cancel button when exporting
                        if (progressBar.IsCancel)
                        {
                            isCancelExport = true;
                            break;
                        }

                        nCount++;
                        dPercent = (nCount / (sprinklers.Count * 1.0)) * 100.0;
                        progressBar.tbxMessage.Text = "Complete : " + dPercent.ToString("0.00") + "% ";
                        progressBar.IncrementProgressBar();
                    }
                }
                catch (Exception)
                { }
                finally
                {
                    progressBar.Dispose();
                }

                tranGr.Assimilate();

                return Result.Succeeded;
            }
            catch (System.Exception)
            { }
            finally
            {
                if (App.m_SprinkerUpForm != null && App.m_SprinkerUpForm.IsDisposed == false)
                {
                    App.m_SprinkerUpForm.Show(App.hWndRevit);
                }
                DisplayService.SetFocus(new HandleRef(null, App.m_SprinkerUpForm.Handle));
            }
            return Result.Cancelled;
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
        public static List<FamilyInstance> SelectSprinklers(bool isD15)
        {
            List<FamilyInstance> list = new List<FamilyInstance>();
            try
            {
                var pickedObjs = Global.UIDoc.Selection.PickObjects(ObjectType.Element, new SprinklerFilter(isD15), "Pick Sprinklers: ");

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

        public static bool cc(Transaction tran, FamilyInstance instance, List<ElementId> selectedIds, bool isConnectNipple, bool isConnectTee, FamilySymbol fmlNipple, double pipeSize, ElementId pipeTypeId, bool isUp)
        {
            double d = UnitUtils.ConvertToInternalUnits(70, DisplayUnitType.DUT_MILLIMETERS);

            bool result = false;

            XYZ direction = isUp ? -XYZ.BasisZ : XYZ.BasisZ;

            //Set d = 25
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
                double radius = 100; //mm
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

                bool isPipe = false;

                List<Element> pipeList = new List<Element>();
                if (collector.GetElementCount() == 0)
                {
                    isPipe = true;
                    pipeList = GetPipes(selectedIds, sprinkle_point);
                }
                else
                    pipeList = collector.ToElements().ToList();

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

                    if (isPipe)
                    {
                        var lineUn = Line.CreateUnbound((curve as Line).Origin, (curve as Line).Direction);
                        var resultPipe = lineUn.Project(sprinkle_point);

                        XYZ pointProject = resultPipe.XYZPoint;
                        XYZ pointProject2d = Common.ToPoint2D(pointProject);
                        var distancePoint2d = pointProject2d.DistanceTo(sprinker2d);
                        if (!Common.IsEqual(distancePoint2d, 0))
                        {
                            if (d < distancePoint2d)
                                return result;
                            else
                            {
                                var vector = pointProject2d - sprinker2d;

                                ElementTransformUtils.MoveElement(Global.UIDoc.Document, instance.Id, vector.Normalize() * pointProject2d.DistanceTo(sprinker2d));

                                Global.UIDoc.Document.Regenerate();

                                sprinkle_point = (instance.Location as LocationPoint).Point;
                            }
                        }
                        else
                        {
                            Line line = Line.CreateBound(p0, pointProject);
                            (pipe.Location as LocationCurve).Curve = line;

                            p0 = line.GetEndPoint(0);
                            p1 = line.GetEndPoint(1);
                        }
                    }
                    else
                    {
                        var resultPipe = curve.Project(sprinkle_point);

                        XYZ pointProject = resultPipe.XYZPoint;
                        XYZ pointProject2d = Common.ToPoint2D(pointProject);

                        var distancePoint2d = pointProject2d.DistanceTo(sprinker2d);
                        if (!Common.IsEqual(distancePoint2d, 0))
                        {
                            if (d < distancePoint2d)
                            {
                                var con = Common.GetConnectorClosestTo(pipe, sprinkle_point);
                                if (con.IsConnected)
                                    return result;
                                else if (!con.IsConnected && !App.m_SprinkerUpForm.isElbow)
                                    return result;
                            }
                            else
                            {
                                var vector = pointProject2d - sprinker2d;

                                ElementTransformUtils.MoveElement(Global.UIDoc.Document, instance.Id, vector.Normalize() * pointProject2d.DistanceTo(sprinker2d));

                                Global.UIDoc.Document.Regenerate();

                                sprinkle_point = (instance.Location as LocationPoint).Point;
                            }
                        }
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

                    pOn = curveExtend3d_temp.Project(pcenter).XYZPoint;

                    (newPipeZ.Location as LocationCurve).Curve = Line.CreateBound(pOn, pcenter);

                    newPipeZ.LookupParameter("Diameter").Set(dFt);

                    var pipeType = Global.UIDoc.Document.GetElement(pipeTypeId) as PipeType;
                    if (pipeType != null)
                        newPipeZ.PipeType = pipeType;

                    Global.UIDoc.Document.Regenerate();

                    try
                    {
                        if (App.m_SprinkerUpForm.isElbow)
                        {
                            if (CheckPipeIsEnd(pipe, sprinkle_point))
                            {
                                if (isConnectNipple)
                                {
                                    var pipeCut = ConnectTeeAndNipple(pipe, newPipeZ, instance, fmlNipple, pOn, pcenter, sprinkle_point);
                                    if (pipeCut != null)
                                        selectedIds.Add(pipeCut.Id);
                                }
                                else
                                {
                                    if (GetPreferredJunctionType(pipe) == PreferredJunctionType.Tee)
                                    {
                                        var curvePipeMain = pipe.GetCurve();

                                        var curvePipeDung = newPipeZ.GetCurve();

                                        Line line = Line.CreateUnbound((curvePipeDung as Line).Origin, (curvePipeDung as Line).Direction);

                                        IntersectionResultArray array;

                                        line.Intersect(curvePipeMain, out array);
                                        var point = array.get_Item(0).XYZPoint;

                                        CreateTeeFitting(pipe, newPipeZ, point, out Pipe pipe1);
                                        if (pipe1 != null)
                                            selectedIds.Add(pipe1.Id);
                                    }
                                    else
                                        se(pipe as MEPCurve, newPipeZ as MEPCurve);

                                    Connector c4 = Common.GetConnectorClosestTo(instance, pcenter);
                                    Connector c5 = Common.GetConnectorClosestTo(newPipeZ, sprinkle_point);

                                    var reducer = Global.UIDoc.Document.Create.NewTransitionFitting(c5, c4);
                                }
                            }
                            else
                            {
                                if (isConnectNipple)
                                    ConnectNipple(pipe, newPipeZ, instance, fmlNipple, pOn, pcenter, sprinkle_point);
                                else
                                {
                                    Connector c4 = Common.GetConnectorClosestTo(instance, pcenter);
                                    Connector c5 = Common.GetConnectorClosestTo(newPipeZ, sprinkle_point);

                                    var reducer = Global.UIDoc.Document.Create.NewTransitionFitting(c5, c4);

                                    Connector c1 = Common.GetConnectorClosestTo(pipe, pOn);

                                    Connector c3 = Common.GetConnectorClosestTo(newPipeZ, pOn);

                                    var elbow = Global.UIDoc.Document.Create.NewElbowFitting(c1, c3);
                                }
                            }
                        }
                        else
                        {
                            if (isConnectNipple)
                            {
                                var pipeCut = ConnectTeeAndNipple(pipe, newPipeZ, instance, fmlNipple, pOn, pcenter, sprinkle_point);
                                if (pipeCut != null)
                                    selectedIds.Add(pipeCut.Id);
                            }
                            else if (isConnectTee)
                            {
                                var pipeCut = ConnectTee(pipe, newPipeZ, instance, pcenter, sprinkle_point);
                                if (pipeCut != null)
                                    selectedIds.Add(pipeCut.Id);
                            }
                            else
                            {
                                if (GetPreferredJunctionType(pipe) == PreferredJunctionType.Tee)
                                {
                                    var curvePipeMain = pipe.GetCurve();

                                    var curvePipeDung = newPipeZ.GetCurve();

                                    Line line = Line.CreateUnbound((curvePipeDung as Line).Origin, (curvePipeDung as Line).Direction);

                                    IntersectionResultArray array;

                                    line.Intersect(curvePipeMain, out array);
                                    var point = array.get_Item(0).XYZPoint;

                                    CreateTeeFitting(pipe, newPipeZ, point, out Pipe pipe1);
                                    if (pipe1 != null)
                                        selectedIds.Add(pipe1.Id);
                                }
                                else
                                    se(pipe as MEPCurve, newPipeZ as MEPCurve);

                                Connector c4 = Common.GetConnectorClosestTo(instance, pcenter);
                                Connector c5 = Common.GetConnectorClosestTo(newPipeZ, sprinkle_point);

                                var reducer = Global.UIDoc.Document.Create.NewTransitionFitting(c5, c4);
                            }
                        }

                        result = true;
                    }
                    catch (System.Exception ex)
                    {
                        tran.RollBack();
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

        public static List<Element> GetPipes(List<ElementId> selectedIds, XYZ sprinklerPoint)
        {
            Dictionary<Pipe, double> keyValuePairs = new Dictionary<Pipe, double>();
            List<Element> pipes = new List<Element>();
            foreach (ElementId elementId in selectedIds)
            {
                var pipePick = Global.UIDoc.Document.GetElement(elementId) as Pipe;
                if (pipePick == null)
                    continue;

                var con = Common.GetConnectorClosestTo(pipePick, sprinklerPoint);
                if (con == null)
                    continue;

                var conOrigin2D = Common.ToPoint2D(con.Origin);
                var sprinklerPoint2D = Common.ToPoint2D(sprinklerPoint);

                var distance = conOrigin2D.DistanceTo(sprinklerPoint2D);

                if (con.IsConnected)
                {
                    double d = UnitUtils.ConvertToInternalUnits(30, DisplayUnitType.DUT_MILLIMETERS);
                    if (distance > d)
                        continue;
                }
                keyValuePairs.Add(pipePick, distance);
            }

            if (keyValuePairs.Count > 0)
            {
                var min = keyValuePairs.Min(x => x.Value);

                var pipeCloset = keyValuePairs.FirstOrDefault(x => x.Value == min);

                if (pipeCloset.Key != null)
                    pipes.Add(pipeCloset.Key);
            }

            return pipes;
        }

        public static bool CheckPipeIsEnd(Pipe pipe, XYZ point)
        {
            var con = Common.GetConnectorClosestTo(pipe, point);

            return con.IsConnected;
        }

        public static FamilyInstance se(MEPCurve mepCurveSplit1, MEPCurve mepCurveSplit2)
        {
            try
            {
                FamilyInstance familyInstance = null;

                var locationCurve1 = mepCurveSplit1.GetCurve();
                var line1 = locationCurve1 as Line;

                var locationCurve2 = mepCurveSplit2.GetCurve();
                var line2 = locationCurve2 as Line;

                var p10 = line2.GetEndPoint(0);
                var p11 = line2.GetEndPoint(1);

                var inter1 = locationCurve1.Project(p10);
                var inter2 = locationCurve1.Project(p11);

                if (inter1 == null || inter2 == null)
                    return null;

                var d1 = inter1.XYZPoint.DistanceTo(p10);
                var d2 = inter2.XYZPoint.DistanceTo(p11);

                if (d1 < d2)
                {
                    var con = Common.GetConnectorClosestTo(mepCurveSplit2, p10);
                    familyInstance = Global.UIDoc.Document.Create.NewTakeoffFitting(con, mepCurveSplit1);
                }
                else
                {
                    var con = Common.GetConnectorClosestTo(mepCurveSplit2, p11);
                    familyInstance = Global.UIDoc.Document.Create.NewTakeoffFitting(con, mepCurveSplit1);
                }

                return familyInstance;
            }
            catch (System.Exception ex)
            {
                return null;
            }
        }

        public static void ConnectNipple(Pipe pipe, Pipe newPipeZ, FamilyInstance instance, FamilySymbol fmlNipple, XYZ pOn, XYZ pcenter, XYZ sprinkle_point)
        {
            var nipple = Global.UIDoc.Document.Create.NewFamilyInstance(Line.CreateBound(pOn, pcenter).Origin, fmlNipple, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
            if (nipple != null)
            {
                var paraDia = newPipeZ.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble();
                nipple.LookupParameter("Nominal Diameter").Set(paraDia);
            }

            Common.RotateLineC2(Global.UIDoc.Document, nipple, Line.CreateBound(pOn, pcenter));

            var connectors = Common.ToList(nipple.MEPModel.ConnectorManager.Connectors);

            Connector c1 = Common.GetConnectorClosestTo(pipe, pOn);

            Connector c3 = Common.GetConnectorClosestTo(newPipeZ, pOn);
            // Get Comector vertical pipe bottom

            Connector cNipple = connectors.OrderBy(x => x.Origin.Z).FirstOrDefault();

            var elbow = Global.UIDoc.Document.Create.NewElbowFitting(c1, c3);

            Connector c4 = Common.GetConnectorClosestTo(instance, pcenter);
            Connector c5 = Common.GetConnectorClosestTo(newPipeZ, sprinkle_point);

            var reducer = Global.UIDoc.Document.Create.NewTransitionFitting(c5, c4);

            //Move Nipple va Reducer

            var conElbow = Common.ToList(elbow.MEPModel.ConnectorManager.Connectors).OrderBy(x => x.Origin.Z).LastOrDefault();
            var vector = conElbow.Origin - cNipple.Origin;

            ElementTransformUtils.MoveElement(Global.UIDoc.Document, nipple.Id, vector.Normalize() * conElbow.Origin.DistanceTo(cNipple.Origin));

            var cReducer = Common.ToList(reducer.MEPModel.ConnectorManager.Connectors).OrderBy(x => x.Origin.Z).FirstOrDefault();
            Connector cNip = connectors.OrderBy(x => x.Origin.Z).LastOrDefault();

            var vectorMoveReducer = cNip.Origin - cReducer.Origin;

            ElementTransformUtils.MoveElement(Global.UIDoc.Document, reducer.Id, vectorMoveReducer.Normalize() * cNip.Origin.DistanceTo(cReducer.Origin));

            Global.UIDoc.Document.Delete(newPipeZ.Id);

            var cConnectElbow = Common.GetConnectorClosestTo(nipple, conElbow.Origin);

            if (cConnectElbow != null)
                conElbow.ConnectTo(cConnectElbow);

            var cConnectReducer = Common.GetConnectorClosestTo(nipple, cReducer.Origin);

            if (cConnectReducer != null)
                cReducer.ConnectTo(cConnectReducer);
        }

        private static PreferredJunctionType GetPreferredJunctionType(Pipe pipe)
        {
            var pipeType = pipe.PipeType as PipeType;

            return pipeType.RoutingPreferenceManager.PreferredJunctionType;
        }

        public static Pipe ConnectTeeAndNipple(Pipe pipe, Pipe newPipeZ, FamilyInstance instance, FamilySymbol fmlNipple, XYZ pOn, XYZ pcenter, XYZ sprinkle_point)
        {
            Pipe pipe1 = null;
            var curvePipeMain = pipe.GetCurve();

            var curvePipeDung = newPipeZ.GetCurve();

            Line line = Line.CreateUnbound((curvePipeDung as Line).Origin, (curvePipeDung as Line).Direction);

            IntersectionResultArray array;

            line.Intersect(curvePipeMain, out array);
            var point = array.get_Item(0).XYZPoint;

            FamilyInstance tee = null;

            if (GetPreferredJunctionType(pipe) == PreferredJunctionType.Tee)
                tee = CreateTeeFitting(pipe, newPipeZ, point, out pipe1);
            else
                tee = se(pipe as MEPCurve, newPipeZ as MEPCurve);

            Connector c4 = Common.GetConnectorClosestTo(instance, pcenter);
            Connector c5 = Common.GetConnectorClosestTo(newPipeZ, sprinkle_point);

            var reducer = Global.UIDoc.Document.Create.NewTransitionFitting(c5, c4);

            var nipple = Global.UIDoc.Document.Create.NewFamilyInstance(Line.CreateBound(pOn, pcenter).Origin, fmlNipple, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
            if (nipple != null)
            {
                var paraDia = newPipeZ.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble();
                nipple.LookupParameter("Nominal Diameter").Set(paraDia);
            }

            Common.RotateLineC2(Global.UIDoc.Document, nipple, Line.CreateBound(pOn, pcenter));

            var connectors = Common.ToList(nipple.MEPModel.ConnectorManager.Connectors);

            Connector cNipple = connectors.OrderBy(x => x.Origin.Z).FirstOrDefault();

            var conElbow = Common.ToList(tee.MEPModel.ConnectorManager.Connectors).OrderBy(x => x.Origin.Z).LastOrDefault();
            var vector = conElbow.Origin - cNipple.Origin;

            ElementTransformUtils.MoveElement(Global.UIDoc.Document, nipple.Id, vector.Normalize() * conElbow.Origin.DistanceTo(cNipple.Origin));

            var cReducer = Common.ToList(reducer.MEPModel.ConnectorManager.Connectors).OrderBy(x => x.Origin.Z).FirstOrDefault();
            Connector cNip = connectors.OrderBy(x => x.Origin.Z).LastOrDefault();

            var vectorMoveReducer = cNip.Origin - cReducer.Origin;

            ElementTransformUtils.MoveElement(Global.UIDoc.Document, reducer.Id, vectorMoveReducer.Normalize() * cNip.Origin.DistanceTo(cReducer.Origin));

            Global.UIDoc.Document.Delete(newPipeZ.Id);

            var cConnectElbow = Common.GetConnectorClosestTo(nipple, conElbow.Origin);

            if (cConnectElbow != null)
                conElbow.ConnectTo(cConnectElbow);

            var cConnectReducer = Common.GetConnectorClosestTo(nipple, cReducer.Origin);

            if (cConnectReducer != null)
                cReducer.ConnectTo(cConnectReducer);

            return pipe1;
        }

        public static Pipe ConnectTee(Pipe pipe, Pipe newPipeZ, FamilyInstance instance, XYZ pcenter, XYZ sprinkle_point)
        {
            Pipe pipe1 = null;
            var curvePipeMain = pipe.GetCurve();

            var curvePipeDung = newPipeZ.GetCurve();

            Line line = Line.CreateUnbound((curvePipeDung as Line).Origin, (curvePipeDung as Line).Direction);

            IntersectionResultArray array;

            line.Intersect(curvePipeMain, out array);
            var point = array.get_Item(0).XYZPoint;

            var cSprinkle = Common.ToList(instance.MEPModel.ConnectorManager.Connectors).FirstOrDefault();

            FamilyInstance fitting = null;

            if (GetPreferredJunctionType(pipe) == PreferredJunctionType.Tee)
                fitting = CreateTeeFitting(pipe, newPipeZ, point, out pipe1);
            else
                fitting = se(pipe as MEPCurve, newPipeZ as MEPCurve);

            Global.UIDoc.Document.Delete(newPipeZ.Id);

            var cTee = Common.ToList(fitting.MEPModel.ConnectorManager.Connectors).OrderBy(x => x.Origin.Z).LastOrDefault();

            var vector = cTee.Origin - cSprinkle.Origin;

            ElementTransformUtils.MoveElement(Global.UIDoc.Document, instance.Id, vector);

            cTee.ConnectTo(cSprinkle);

            return pipe1;
        }

        public static FamilyInstance CreateTeeFitting(Pipe pipeMain, Pipe pipeCurrent, XYZ splitPoint, out Pipe main2)
        {
            Line line1 = null;
            Line line2 = null;
            var curve = (pipeMain.Location as LocationCurve).Curve;

            var p0 = curve.GetEndPoint(0);
            var p1 = curve.GetEndPoint(1);

            var pipeTempMain1 = pipeMain;

            if ((curve as Line).Direction.IsAlmostEqualTo(XYZ.BasisX))
            {
                line2 = Line.CreateBound(p0, splitPoint);
                line1 = Line.CreateBound(splitPoint, p1);
            }
            else
            {
                line1 = Line.CreateBound(p0, splitPoint);
                line2 = Line.CreateBound(splitPoint, p1);
            }

            (pipeTempMain1.Location as LocationCurve).Curve = line1;

            var newPlace = new XYZ(0, 0, 0);
            var elemIds = ElementTransformUtils.CopyElement(
               Global.UIDoc.Document, pipeTempMain1.Id, newPlace);

            main2 = Global.UIDoc.Document.GetElement(elemIds.ToList()[0]) as Pipe;

            (main2.Location as LocationCurve).Curve = line2;

            //Connect
            var c3 = Common.GetConnectorClosestTo(pipeTempMain1, splitPoint);
            var c4 = Common.GetConnectorClosestTo(main2, splitPoint);
            var c5 = Common.GetConnectorClosestTo(pipeCurrent, splitPoint);

            var fitting = Global.UIDoc.Document.Create.NewTeeFitting(c3, c4, c5);

            return fitting;
        }

        public static FamilyInstance CreateTeeFittingSprinkler(Pipe pipeMain, FamilyInstance instance, XYZ splitPoint, out Pipe main2)
        {
            var curve = (pipeMain.Location as LocationCurve).Curve;

            var p0 = curve.GetEndPoint(0);
            var p1 = curve.GetEndPoint(1);

            var pipeTempMain1 = pipeMain;

            var line1 = Line.CreateBound(p0, splitPoint);
            (pipeTempMain1.Location as LocationCurve).Curve = line1;

            var newPlace = new XYZ(0, 0, 0);
            var elemIds = ElementTransformUtils.CopyElement(
               Global.UIDoc.Document, pipeTempMain1.Id, newPlace);

            main2 = Global.UIDoc.Document.GetElement(elemIds.ToList()[0]) as Pipe;

            var line2 = Line.CreateBound(splitPoint, p1);
            (main2.Location as LocationCurve).Curve = line2;

            //Connect
            var c3 = Common.GetConnectorClosestTo(pipeTempMain1, splitPoint);
            var c4 = Common.GetConnectorClosestTo(main2, splitPoint);
            var c5 = Common.GetConnectorClosestTo(instance, splitPoint);

            try
            {
                var fitting = Global.UIDoc.Document.Create.NewTeeFitting(c3, c4, c5);

                return fitting;
            }
            catch (System.Exception ex)
            {
                return null;
            }
        }
    }

    public class SprinklerFilter : ISelectionFilter
    {
        public bool _isD15 = false;

        public SprinklerFilter(bool isD15)
        {
            _isD15 = isD15;
        }

        public bool AllowElement(Element element)
        {
            if (_isD15)
            {
                if (element != null
                    && element.Category != null
                    && (BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_Sprinklers
                    && element.GetTypeId() != null
                    && Global.UIDoc.Document.GetElement(element.GetTypeId()) != null
                    && Global.UIDoc.Document.GetElement(element.GetTypeId()).LookupParameter("Diameter") != null
                    && Global.UIDoc.Document.GetElement(element.GetTypeId()).LookupParameter("Diameter").AsValueString().Contains("15"))
                    return true;
            }
            else
            {
                if (element != null
                                    && element.Category != null
                                    && (BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_Sprinklers
                                    && element.GetTypeId() != null
                                    && Global.UIDoc.Document.GetElement(element.GetTypeId()) != null
                                    && Global.UIDoc.Document.GetElement(element.GetTypeId()).LookupParameter("Diameter") != null
                                    && Global.UIDoc.Document.GetElement(element.GetTypeId()).LookupParameter("Diameter").AsValueString().Contains("20"))
                    return true;
            }

            return false;
        }

        public bool AllowReference(Reference refer, XYZ point)
        {
            return false;
        }
    }
}