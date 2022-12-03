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
    public class Cmd2LevelSmart : IExternalCommand
    {
        private List<ElementId> _MainPipes = new List<ElementId>();

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //if (App.CROK() != true)
            //    return Result.Failed;

            Global.UIApp = commandData.Application;
            Global.RVTApp = commandData.Application.Application;
            Global.UIDoc = commandData.Application.ActiveUIDocument;
            Global.AppCreation = commandData.Application.Application.Create;

            _MainPipes.Clear();

            //Show form
            if (App.Show2LevelSmartForm() == false)
                return Result.Cancelled;

            //Pick pipe chinh
            //var pipe_main = PickPipe();
            List<Pipe> pipe_mains = Cmd2LevelSmart.PickPipe();

            if (pipe_mains == null || pipe_mains.Count == 0)
                return Result.Cancelled;

            List<Pipe> pipes = Cmd2LevelSmart.PickPipe();
            if (pipes == null || pipes.Count == 0)
                return Result.Cancelled;

            pipes = (from Pipe p in pipes
                     where /*p.Id != pipe_main.Id*/ pipe_mains.Find(item => item.Id == p.Id) == null
                     select p).ToList();

            var pairs = pp(pipes);
            if (pairs == null || pairs.Count == 0)
                return Result.Cancelled;

            //_MainPipes.Add(pipe_main.Id);

            var ids = (from Pipe p in pipe_mains
                       where p.Id != ElementId.InvalidElementId
                       select p.Id).ToList();

            _MainPipes.AddRange(ids);

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

        private List<pp> pp(List<Pipe> pipes)
        {
            //             if (pipes.Count < 2)
            //                 return null;

            double temp = 100;

            List<pp> pairs = new List<pp>();

            foreach (Pipe pipe in pipes)
            {
                var curve_sub = pipe.GetCurve();

                var line = Line.CreateBound(curve_sub.GetEndPoint(0), curve_sub.GetEndPoint(1));

                //Expand
                var p0_ex = line.Evaluate(line.GetEndParameter(0) - temp, false);
                var p1_ex = line.Evaluate(line.GetEndParameter(1) + temp, false);

                var expand = Line.CreateBound(p0_ex, p1_ex);

                foreach (Pipe pipe2 in pipes)
                {
                    if (pipe.Id == pipe2.Id)
                        continue;

                    var curve_sub2 = pipe2.GetCurve();
                    IntersectionResultArray arr = new IntersectionResultArray();
                    var result = curve_sub2.Intersect(expand, out arr);

                    if (result == SetComparisonResult.Equal)
                    {
                        //Dong tam

                        var find = pairs.Find(item => ce(item, pipe, pipe2));
                        if (find == null)
                        {
                            var newPair = new pp(pipe, pipe2);
                            pairs.Add(newPair);
                        }
                    }
                }
            }

            List<ElementId> allAddeds1 = (from pp pair in pairs
                                          where pair._Pipe1 != null
                                          select pair._Pipe1.Id).ToList();

            List<ElementId> allAddeds2 = (from pp pair in pairs
                                          where pair._Pipe2 != null
                                          select pair._Pipe2.Id).ToList();

            List<ElementId> rest = (from Pipe pipe in pipes
                                    where allAddeds1.Contains(pipe.Id) == false && allAddeds2.Contains(pipe.Id) == false
                                    select pipe.Id).ToList();

            foreach (ElementId id in rest)
            {
                var pipe = Global.UIDoc.Document.GetElement(id) as Pipe;
                if (pipe == null)
                    continue;

                var newPair = new pp(pipe, null);
                pairs.Add(newPair);
            }

            return pairs;
        }

        private bool ce(pp pair, Pipe pipe1, Pipe pipe2)
        {
            if (pair._Pipe1.Id == pipe1.Id && pair._Pipe2.Id == pipe2.Id)
                return true;
            if (pair._Pipe1.Id == pipe2.Id && pair._Pipe2.Id == pipe1.Id)
                return true;

            return false;
        }

        private bool ps(Pipe pipe1, Pipe pipe2, double diameter, PipeType typePipe)
        {
            SubTransaction sub = new SubTransaction(Global.UIDoc.Document);
            sub.Start();

            try
            {
                XYZ inter_main1 = null;
                XYZ p0_sub1 = null;
                XYZ p1_sub1 = null;
                XYZ inters_sub1 = null;

                bool twoPipes = pipe2 != null ? true : false;

                Pipe main = null;

                foreach (ElementId pId in _MainPipes)
                {
                    var m = Global.UIDoc.Document.GetElement(pId) as Pipe;
                    if (dv(m, pipe1, twoPipes, out inter_main1, out p0_sub1, out p1_sub1, out inters_sub1) == true)
                    {
                        main = m;
                        break;
                    }
                }

                // Common.CreateModelLine(p1_sub1, Common.NewPoint(p1_sub1));

                if (pipe2 == null)
                {
                    (pipe1.Location as LocationCurve).Curve = Line.CreateBound(p0_sub1, inters_sub1);

                    pipe2 = Common.Clone(pipe1) as Pipe;
                    (pipe2.Location as LocationCurve).Curve = Line.CreateBound(inters_sub1, p1_sub1);

                    //Set lai gia tri
                    p1_sub1 = inters_sub1;
                }

                XYZ inter_main2 = null;
                XYZ p0_sub2 = null;
                XYZ p1_sub2 = null;
                XYZ inters_sub2 = null;

                if (dv(main, pipe2, true, out inter_main2, out p0_sub2, out p1_sub2, out inters_sub2) == false)
                {
                    sub.RollBack();
                    return false;
                }
                //Common.CreateModelLine(p1_sub2, Common.NewPoint(p1_sub2));

                if (inter_main1.DistanceTo(inter_main2) > 0.001)
                {
                    sub.RollBack();
                    return false; //Check
                }

                if (inters_sub1.DistanceTo(inters_sub2) > 0.001)
                {
                    sub.RollBack();
                    return false; //Check
                }

                //Create vertical pipe
                var vertical = Common.Clone(pipe1) as Pipe;
                vertical.PipeType = typePipe;

                vertical.LookupParameter("Diameter").Set(diameter);

                (vertical.Location as LocationCurve).Curve = Line.CreateBound(inter_main1, inters_sub1);

                //Conect to main pipe
                if (gr(main) == PreferredJunctionType.Tee)
                {
                    //Try
                    if (main as Pipe != null && vertical as Pipe != null && inter_main1 != null)
                    {
                        Pipe main2 = null;

                        var fitting = ee(main as Pipe, vertical as Pipe, inter_main1, out main2);
                        if (main2 != null && fitting != null)
                        {
                            _MainPipes.Add(main2.Id);
                        }
                        else
                        {
                            sub.RollBack();
                            return false;
                        }
                    }
                }
                else
                {
                    if (CreateFittingForMEPUtils.se(main, vertical) == false)
                    {
                        sub.RollBack();
                        return false;
                    }
                }

                //Process for sub pipes
                var c5 = Common.GetConnectorClosestTo(vertical, inters_sub1);

                double dMoi = 15 * Common.mmToFT;

                double d10 = 10 * Common.mmToFT;
                double dtemp = 2;

                if (g(diameter, pipe1.Diameter) && g(diameter, pipe2.Diameter))
                {
                    //Connect to sub
                    (pipe1.Location as LocationCurve).Curve = Line.CreateBound(p0_sub1, p1_sub1);
                    (pipe2.Location as LocationCurve).Curve = Line.CreateBound(p0_sub2, p1_sub2);

                    //Connect
                    var c3 = Common.GetConnectorClosestTo(pipe1, inters_sub1);
                    var c4 = Common.GetConnectorClosestTo(pipe2, inters_sub1);

                    if (CreateFittingForMEPUtils.CreatTee(c3, c4, c5) == null)
                    {
                        sub.RollBack();
                        return false;
                    }
                }
                else
                {
                    Connector c3 = null;
                    Pipe pipe_moi_1 = null;

                    XYZ v = null;
                    if (g(diameter, pipe1.Diameter) == false)
                    {
                        //Tao mot ong mồi
                        pipe_moi_1 = Common.Clone(vertical) as Pipe;

                        var line = Line.CreateBound(p0_sub1, p1_sub1);

                        var p1 = p0_sub1;// line.Evaluate(dtemp, false);

                        if (p0_sub1.DistanceTo(inters_sub1) < 0.01)
                        {
                            p1 = p1_sub1;
                        }

                        (pipe_moi_1.Location as LocationCurve).Curve = Line.CreateBound(inters_sub1, p1);

                        c3 = Common.GetConnectorClosestTo(pipe_moi_1, inters_sub1);

                        v = line.Direction;
                    }
                    else
                    {
                        //Connect to sub

                        var line = Line.CreateBound(p0_sub1, p1_sub1);
                        (pipe1.Location as LocationCurve).Curve = line;
                        c3 = Common.GetConnectorClosestTo(pipe1, inters_sub1);

                        v = line.Direction;
                    }

                    Connector c4 = null;
                    Pipe pipe_moi_2 = null;

                    if (g(diameter, pipe2.Diameter) == false)
                    {
                        //Tao mot ong mồi
                        pipe_moi_2 = Common.Clone(vertical) as Pipe;

                        var line = Line.CreateBound(p0_sub2, p1_sub2);

                        var p1 = p0_sub2;// line.Evaluate(dtemp, false);
                        if (p0_sub2.DistanceTo(inters_sub2) < 0.01)
                        {
                            p1 = p1_sub2;
                        }

                        (pipe_moi_2.Location as LocationCurve).Curve = Line.CreateBound(inters_sub2, p1);

                        c4 = Common.GetConnectorClosestTo(pipe_moi_2, inters_sub2);
                    }
                    else
                    {
                        (pipe2.Location as LocationCurve).Curve = Line.CreateBound(p0_sub2, p1_sub2);
                        c4 = Common.GetConnectorClosestTo(pipe2, inters_sub1);
                    }

                    //Connect
                    FamilyInstance fitting = CreateFittingForMEPUtils.CreatTee(c3, c4, c5);
                    if (fitting == null)
                    {
                        sub.RollBack/*Commit*/();
                        return false;
                    }

                    if (fitting != null)
                    {
                        Connector main1 = null;
                        Connector main2 = null;
                        Connector tee = null;
                        Common.GetInfo(fitting, v, out main1, out main2, out tee);

                        var mep1 = ee(main1.AllRefs);
                        var mep2 = ee(main2.AllRefs);

                        var main1_p = main1.Origin;
                        var v_m_1 = main1.CoordinateSystem.BasisZ;

                        var main2_p = main2.Origin;
                        var v_m_2 = main2.CoordinateSystem.BasisZ;

                        //1
                        if (pipe_moi_1 != null)
                        {
                            if (mep1 != null && pipe_moi_1.Id == mep1.Id)
                            {
                                var line = Line.CreateUnbound(main1_p, v_m_1 * 10);

                                var p = line.Evaluate(d10, false);
                                (pipe_moi_1.Location as LocationCurve).Curve = Line.CreateBound(main1_p, p);

                                var futher = inters_sub1.DistanceTo(p0_sub1) > inters_sub1.DistanceTo(p1_sub1) ? p0_sub1 : p1_sub1;

                                ll(pipe1, futher, p);

                                CreateFittingForMEPUtils.CreatTransitionFitting(pipe1, pipe_moi_1);
                            }
                            else if (mep2 != null && pipe_moi_1.Id == mep2.Id)
                            {
                                var line = Line.CreateUnbound(main2_p, v_m_2 * 10);

                                var p = line.Evaluate(d10, false);
                                (pipe_moi_1.Location as LocationCurve).Curve = Line.CreateBound(main2_p, p);

                                var futher = inters_sub1.DistanceTo(p0_sub1) > inters_sub1.DistanceTo(p1_sub1) ? p0_sub1 : p1_sub1;

                                ll(pipe1, futher, p);

                                CreateFittingForMEPUtils.CreatTransitionFitting(pipe1, pipe_moi_1);
                            }
                        }

                        //2
                        if (pipe_moi_2 != null)
                        {
                            if (mep1 != null && pipe_moi_2.Id == mep1.Id)
                            {
                                var line = Line.CreateUnbound(main1_p, v_m_1 * 10);

                                var p = line.Evaluate(d10, false);
                                (pipe_moi_2.Location as LocationCurve).Curve = Line.CreateBound(main1_p, p);

                                var futher = inters_sub2.DistanceTo(p0_sub2) > inters_sub2.DistanceTo(p1_sub2) ? p0_sub2 : p1_sub2;

                                ll(pipe2, futher, p);

                                CreateFittingForMEPUtils.CreatTransitionFitting(pipe2, pipe_moi_2);
                            }
                            else if (mep2 != null && pipe_moi_2.Id == mep2.Id)
                            {
                                var line = Line.CreateUnbound(main2_p, v_m_2 * 10);

                                var p = line.Evaluate(d10, false);
                                (pipe_moi_2.Location as LocationCurve).Curve = Line.CreateBound(main2_p, p);

                                var futher = inters_sub2.DistanceTo(p0_sub2) > inters_sub2.DistanceTo(p1_sub2) ? p0_sub2 : p1_sub2;

                                ll(pipe2, futher, p);

                                CreateFittingForMEPUtils.CreatTransitionFitting(pipe2, pipe_moi_2);
                            }
                        }
                    }
                }
                sub.Commit();
                return true;
            }
            catch (System.Exception ex)
            {
                sub.RollBack();
                return false;
            }
        }

        public static FamilyInstance ee(Pipe pipeMain, Pipe pipeCurrent, XYZ splitPoint, out Pipe main2)
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
            var c5 = Common.GetConnectorClosestTo(pipeCurrent, splitPoint);

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

        private void ll(Pipe pipe, XYZ pOn, XYZ pOther)
        {
            var p0 = pipe.GetCurve().GetEndPoint(0);
            var p1 = pipe.GetCurve().GetEndPoint(1);

            if (p0.DistanceTo(pOn) < p1.DistanceTo(pOn))
            {
                (pipe.Location as LocationCurve).Curve = Line.CreateBound(pOn, pOther);
            }
            else
            {
                (pipe.Location as LocationCurve).Curve = Line.CreateBound(pOther, pOn);
            }
        }

        private Pipe ee(ConnectorSet cs)
        {
            foreach (Connector c in cs)
            {
                Element e = c.Owner;

                if (null != e && e as Pipe != null)
                {
                    return e as Pipe;
                }
            }

            return null;
        }

        private bool g(double d1, double d2)
        {
            if (Math.Abs(d1 - d2) < 0.001)
                return true;

            return false;
        }

        private PreferredJunctionType gr(Pipe pipe)
        {
            var pipeType = pipe.PipeType as PipeType;

            return pipeType.RoutingPreferenceManager.PreferredJunctionType;
        }

        private bool dv(Pipe main, Pipe pipeSub, bool twoPipes, out XYZ inters_main, out XYZ p0_sub, out XYZ p1_sub, out XYZ inters_sub)
        {
            inters_main = null;
            p0_sub = null;
            p1_sub = null;
            inters_sub = null;

            var curve_main = main.GetCurve();
            var curve_main_2d = Line.CreateBound(Common.ToPoint2D(curve_main.GetEndPoint(0)), Common.ToPoint2D(curve_main.GetEndPoint(1)));

            var curve_sub = pipeSub.GetCurve();

            Curve curve2d = null;
            Curve curve3d = null;

            if (twoPipes)
            {
                //Expand
                double t = 10;
                var p0_ex1 = curve_sub.Evaluate(curve_sub.GetEndParameter(0) - t, false);
                var p1_ex2 = curve_sub.Evaluate(curve_sub.GetEndParameter(1) + t, false);
                curve3d = Line.CreateBound(p0_ex1, p1_ex2);

                var expand_2d = Line.CreateBound(Common.ToPoint2D(p0_ex1), Common.ToPoint2D(p1_ex2));

                curve2d = expand_2d;
            }
            else
            {
                curve2d = Line.CreateBound(Common.ToPoint2D(curve_sub.GetEndPoint(0)), Common.ToPoint2D(curve_sub.GetEndPoint(1)));
                curve3d = curve_sub;
            }

            //Check intersection
            IntersectionResultArray arr = new IntersectionResultArray();
            var result = curve_main_2d.Intersect(curve2d, out arr);
            if (result != SetComparisonResult.Overlap)
                return false;

            var pInter_2d = arr.get_Item(0).XYZPoint;

            //Find 3d
            double temp = 200;

            var lineZ = Line.CreateBound(new XYZ(pInter_2d.X, pInter_2d.Y, pInter_2d.Z - temp), new XYZ(pInter_2d.X, pInter_2d.Y, pInter_2d.Z + temp));

            arr = new IntersectionResultArray();
            result = lineZ.Intersect(curve_main, out arr);
            if (result != SetComparisonResult.Overlap)
                return false;

            inters_main = arr.get_Item(0).XYZPoint;

            //Find 3d on two sub pipe
            arr = new IntersectionResultArray();
            result = lineZ.Intersect(curve3d, out arr);
            if (result != SetComparisonResult.Overlap)
                return false;

            inters_sub = arr.get_Item(0).XYZPoint;

            if (twoPipes)
            {
                //Get father point
                int index = inters_sub.DistanceTo(curve_sub.GetEndPoint(0)) > inters_sub.DistanceTo(curve_sub.GetEndPoint(1)) ? 0 : 1;

                if (index == 0)
                {
                    p0_sub = curve_sub.GetEndPoint(index);
                    p1_sub = inters_sub;
                }
                else
                {
                    p0_sub = inters_sub;
                    p1_sub = curve_sub.GetEndPoint(index);
                }
            }
            else
            {
                p0_sub = curve_sub.GetEndPoint(0);
                p1_sub = curve_sub.GetEndPoint(1);
            }

            return true;
        }
    }

    public class pp
    {
        public Pipe _Pipe1 = null;
        public Pipe _Pipe2 = null;

        public pp(Pipe p1, Pipe p2)
        {
            _Pipe1 = p1;
            _Pipe2 = p2;
        }
    }
}