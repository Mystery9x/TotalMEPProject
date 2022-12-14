using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;
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
        private static List<ElementId> m_mainPipes = new List<ElementId>();

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //if (App.CROK() != true)
            //    return Result.Failed;

            Global.UIApp = commandData.Application;
            Global.RVTApp = commandData.Application.Application;
            Global.UIDoc = commandData.Application.ActiveUIDocument;
            Global.AppCreation = commandData.Application.Application.Create;

            //Show form
            if (App.Show2LevelSmartForm() == false)
                return Result.Cancelled;

            return Result.Succeeded;
        }

        public static Result Process()
        {
            m_mainPipes.Clear();
            var diameter = App._2LevelSmartForm.PipeSize * Common.mmToFT;

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

            var pairs = PairPipe(pipes);
            if (pairs == null || pairs.Count == 0)
                return Result.Cancelled;

            //m_mainPipes.Add(pipe_main.Id);

            var ids = (from Pipe p in pipe_mains
                       where p.Id != ElementId.InvalidElementId
                       select p.Id).ToList();

            m_mainPipes.AddRange(ids);

            Transaction t = new Transaction(Global.UIDoc.Document, "Two Level Smart");
            t.Start();

            var pipeType = Global.UIDoc.Document.GetElement(App._2LevelSmartForm.FamilyType) as PipeType;

            foreach (PairPipes pair in pairs)
            {
                var pipe1 = pair._Pipe1;
                var pipe2 = pair._Pipe2;

                HandlerWithPipes(pipe1, pipe2, diameter, pipeType);
            }

            t.Commit();

            return Result.Succeeded;
        }

        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        /// <summary>
        /// Handler With Pipes
        /// </summary>
        /// <param name="firstPipe"></param>
        /// <param name="secondPipe"></param>
        /// <param name="diameter"></param>
        /// <param name="typePipe"></param>
        /// <returns></returns>
        private static bool HandlerWithPipes(Pipe firstPipe, Pipe secondPipe, double diameter, PipeType pipeType)
        {
            try
            {
                using (SubTransaction reSubTrans = new SubTransaction(Global.UIDoc.Document))
                {
                    try
                    {
                        reSubTrans.Start();

                        // Get valid point
                        XYZ intsMain1_pnt = null;
                        XYZ sub1_pnt0 = null;
                        XYZ sub1_pnt1 = null;
                        XYZ sub1Ints_pnt = null;

                        // Check has two pipe valid
                        bool flagTwoPipe = secondPipe != null ? true : false;

                        // Get main pipe
                        Pipe mainPipe = null;
                        if (m_mainPipes == null || m_mainPipes.Count == 0)
                            return false;

                        foreach (ElementId eId in m_mainPipes)
                        {
                            Pipe pipeLoop = Global.UIDoc.Document.GetElement(eId) as Pipe;
                            if (pipeLoop == null)
                                continue;

                            if (HandlerDivide(pipeLoop, firstPipe, flagTwoPipe, out intsMain1_pnt, out sub1_pnt0, out sub1_pnt1, out sub1Ints_pnt) == true)
                            {
                                mainPipe = pipeLoop;
                                break;
                            }
                        }

                        // If only 1 pipe
                        if (secondPipe == null)
                        {
                            (firstPipe.Location as LocationCurve).Curve = Line.CreateBound(sub1_pnt0, sub1Ints_pnt);

                            secondPipe = Common.Clone(firstPipe) as Pipe;
                            (secondPipe.Location as LocationCurve).Curve = Line.CreateBound(sub1Ints_pnt, sub1_pnt1);

                            // Set point value again
                            sub1_pnt1 = sub1Ints_pnt;
                        }

                        // Get valid point
                        XYZ intsMain2_pnt = null;
                        XYZ sub2_pnt0 = null;
                        XYZ sub2_pnt1 = null;
                        XYZ sub2Ints_pnt = null;

                        if (HandlerDivide(mainPipe, secondPipe, true, out intsMain2_pnt, out sub2_pnt0, out sub2_pnt1, out sub2Ints_pnt) == false)
                        {
                            reSubTrans.RollBack();
                            return false;
                        }

                        if (intsMain1_pnt.DistanceTo(intsMain1_pnt) > 0.001)
                        {
                            reSubTrans.RollBack();
                            return false;
                        }

                        if (sub1Ints_pnt.DistanceTo(sub2Ints_pnt) > 0.001)
                        {
                            reSubTrans.RollBack();
                            return false;
                        }

                        // Create vertical pipe
                        Pipe verticalPipe = Common.Clone(firstPipe) as Pipe;
                        verticalPipe.PipeType = pipeType;
                        verticalPipe.LookupParameter("Diameter").Set(diameter);
                        (verticalPipe.Location as LocationCurve).Curve = Line.CreateBound(intsMain1_pnt, sub1Ints_pnt);

                        // Connect vertical pipe with main pipe
                        FamilyInstance bottomTee = null;

                        if (GetPreferredJunctionType(mainPipe) == PreferredJunctionType.Tee)
                        {
                            if (mainPipe as Pipe != null && verticalPipe as Pipe != null && intsMain1_pnt != null)
                            {
                                Pipe main2 = null;

                                bottomTee = CreateTeeFitting(mainPipe as Pipe, verticalPipe as Pipe, intsMain1_pnt, out main2);
                                if (main2 != null && bottomTee != null)
                                {
                                    m_mainPipes.Add(main2.Id);
                                }
                                else
                                {
                                    reSubTrans.RollBack();
                                    return false;
                                }
                            }
                        }
                        else
                        {
                            if (se(mainPipe, verticalPipe) == false)
                            {
                                reSubTrans.RollBack();
                                return false;
                            }
                        }

                        // Process with sub pipes
                        FamilyInstance topTee = null;
                        Connector cntBottom_topTee = GetConnectorClosestTo(verticalPipe, sub1Ints_pnt);

                        double diameter_primer = 15 * Common.mmToFT;

                        double diameter_10 = 10 * Common.mmToFT;

                        // If diameter vertical pipe = diameter main pipe
                        if (g(diameter, firstPipe.Diameter) && g(diameter, secondPipe.Diameter))
                        {
                            //Connect to sub
                            (firstPipe.Location as LocationCurve).Curve = Line.CreateBound(sub1_pnt0, sub1_pnt1);
                            (secondPipe.Location as LocationCurve).Curve = Line.CreateBound(sub2_pnt0, sub2_pnt1);

                            //Connect
                            Connector c3 = GetConnectorClosestTo(firstPipe, sub1Ints_pnt);
                            Connector c4 = GetConnectorClosestTo(secondPipe, sub2Ints_pnt);

                            if (CreateFittingForMEPUtils.CreatTee(c3, c4, cntBottom_topTee) == null)
                            {
                                reSubTrans.RollBack();
                                return false;
                            }
                        }
                        // If diameter vertical pipe < diameter main pipe
                        else
                        {
                            Connector c3 = null;
                            // Create primer pipe 1
                            Pipe primerPipe_1 = null;

                            XYZ tempVector = null;

                            if (g(diameter, firstPipe.Diameter) == false)
                            {
                                // Create primer pipe 1
                                primerPipe_1 = Common.Clone(verticalPipe) as Pipe;
                                Line newLocationCurve = Line.CreateBound(sub1_pnt0, sub1_pnt1);

                                XYZ p1 = sub1_pnt0;

                                if (sub1_pnt0.DistanceTo(sub1Ints_pnt) < 0.01)
                                {
                                    p1 = sub1_pnt1;
                                }

                                (primerPipe_1.Location as LocationCurve).Curve = Line.CreateBound(sub1Ints_pnt, p1);

                                c3 = Common.GetConnectorClosestTo(primerPipe_1, sub1Ints_pnt);

                                tempVector = newLocationCurve.Direction;
                            }
                            else
                            {
                                // Connect to sub pipe
                                var line = Line.CreateBound(sub1_pnt0, sub1_pnt1);
                                (firstPipe.Location as LocationCurve).Curve = line;
                                c3 = Common.GetConnectorClosestTo(firstPipe, sub1Ints_pnt);

                                tempVector = line.Direction;
                            }

                            Connector c4 = null;
                            // Create primer pipe 2
                            Pipe primerPipe_2 = null;

                            if (g(diameter, secondPipe.Diameter) == false)
                            {
                                // Create primer pipe 1
                                primerPipe_2 = Common.Clone(verticalPipe) as Pipe;
                                Line newLocationCurve = Line.CreateBound(sub2_pnt0, sub2_pnt1);

                                XYZ p1 = sub2_pnt0;

                                if (sub2_pnt0.DistanceTo(sub2Ints_pnt) < 0.01)
                                {
                                    p1 = sub2_pnt1;
                                }

                                (primerPipe_2.Location as LocationCurve).Curve = Line.CreateBound(sub2Ints_pnt, p1);

                                c4 = Common.GetConnectorClosestTo(primerPipe_2, sub2Ints_pnt);
                            }
                            else
                            {
                                // Connect to sub pipe
                                var line = Line.CreateBound(sub2_pnt0, sub2_pnt1);
                                (secondPipe.Location as LocationCurve).Curve = line;
                                c4 = Common.GetConnectorClosestTo(firstPipe, sub1Ints_pnt);

                                tempVector = line.Direction;
                            }

                            // Create top tee
                            topTee = CreateFittingForMEPUtils.CreatTee(c3, c4, cntBottom_topTee);

                            if (topTee == null)
                            {
                                reSubTrans.RollBack();
                                return false;
                            }

                            Connector cntTopTee_main1 = null;
                            Connector cntTopTee_main2 = null;
                            Connector cntTopTee_bottom = null;
                            Common.GetInfo(topTee, tempVector, out cntTopTee_main1, out cntTopTee_main2, out cntTopTee_bottom);

                            Pipe checkPrimerTemp1 = ee(cntTopTee_main1.AllRefs);
                            Pipe checkPrimerTemp2 = ee(cntTopTee_main2.AllRefs);

                            XYZ main1_p = cntTopTee_main1.Origin;
                            XYZ v_m_1 = cntTopTee_main1.CoordinateSystem.BasisZ;

                            var main2_p = cntTopTee_main2.Origin;
                            var v_m_2 = cntTopTee_main2.CoordinateSystem.BasisZ;

                            FamilyInstance reducer_1 = null;
                            FamilyInstance reducer_2 = null;

                            // Pipe 1
                            HandlerConnectAccessory(App._2LevelSmartForm.OptionAddNipple,
                                                    firstPipe,
                                                    primerPipe_1,
                                                    checkPrimerTemp1,
                                                    checkPrimerTemp2,
                                                    main1_p,
                                                    v_m_1,
                                                    main2_p,
                                                    v_m_2,
                                                    diameter_10,
                                                    sub1Ints_pnt,
                                                    sub1_pnt0,
                                                    sub1_pnt1,
                                                    topTee,
                                                    out reducer_1);

                            // Pipe 2
                            HandlerConnectAccessory(App._2LevelSmartForm.OptionAddNipple,
                                                    secondPipe,
                                                    primerPipe_2,
                                                    checkPrimerTemp1,
                                                    checkPrimerTemp2,
                                                    main1_p,
                                                    v_m_1,
                                                    main2_p,
                                                    v_m_2,
                                                    diameter_10,
                                                    sub2Ints_pnt,
                                                    sub2_pnt0,
                                                    sub2_pnt1,
                                                    topTee,
                                                    out reducer_2);

                            // Nipple Family
                            HandlerConnectAccessoryNipple(App._2LevelSmartForm.OptionAddNipple, primerPipe_1, reducer_1, topTee);
                            HandlerConnectAccessoryNipple(App._2LevelSmartForm.OptionAddNipple, primerPipe_2, reducer_2, topTee);
                        }

                        reSubTrans.Commit();
                    }
                    catch (Exception)
                    {
                        reSubTrans.RollBack();
                        return false;
                    }
                }
            }
            catch (Exception)
            { }
            return false;
        }

        /// <summary>
        /// Handler Divide
        /// </summary>
        /// <param name="mainPipe"></param>
        /// <param name="subPipe"></param>
        /// <param name="flagTwoPipe"></param>
        /// <param name="intsMain_pnt"></param>
        /// <param name="sub_pnt0"></param>
        /// <param name="sub_pnt1"></param>
        /// <param name="subInts_pnt"></param>
        /// <returns></returns>
        private static bool HandlerDivide(Pipe mainPipe, Pipe subPipe, bool flagTwoPipe, out XYZ intsMain_pnt, out XYZ sub_pnt0, out XYZ sub_pnt1, out XYZ subInts_pnt)
        {
            intsMain_pnt = null;
            sub_pnt0 = null;
            sub_pnt1 = null;
            subInts_pnt = null;

            var curve_main = mainPipe.GetCurve();
            var curve_main_2d = Line.CreateBound(Common.ToPoint2D(curve_main.GetEndPoint(0)), Common.ToPoint2D(curve_main.GetEndPoint(1)));

            var curve_sub = subPipe.GetCurve();

            Curve curve2d = null;
            Curve curve3d = null;

            if (flagTwoPipe)
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

            intsMain_pnt = arr.get_Item(0).XYZPoint;

            //Find 3d on two sub pipe
            arr = new IntersectionResultArray();
            result = lineZ.Intersect(curve3d, out arr);
            if (result != SetComparisonResult.Overlap)
                return false;

            subInts_pnt = arr.get_Item(0).XYZPoint;

            if (flagTwoPipe)
            {
                //Get father point
                int index = subInts_pnt.DistanceTo(curve_sub.GetEndPoint(0)) > subInts_pnt.DistanceTo(curve_sub.GetEndPoint(1)) ? 0 : 1;

                if (index == 0)
                {
                    sub_pnt0 = curve_sub.GetEndPoint(index);
                    sub_pnt1 = subInts_pnt;
                }
                else
                {
                    sub_pnt0 = subInts_pnt;
                    sub_pnt1 = curve_sub.GetEndPoint(index);
                }
            }
            else
            {
                sub_pnt0 = curve_sub.GetEndPoint(0);
                sub_pnt1 = curve_sub.GetEndPoint(1);
            }

            return true;
        }

        private static PreferredJunctionType GetPreferredJunctionType(Pipe pipe)
        {
            var pipeType = pipe.PipeType as PipeType;

            return pipeType.RoutingPreferenceManager.PreferredJunctionType;
        }

        public static FamilyInstance CreateTeeFitting(Pipe pipeMain, Pipe pipeCurrent, XYZ splitPoint, out Pipe main2)
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
            var c3 = GetConnectorClosestTo(pipeTempMain1, splitPoint);
            var c4 = GetConnectorClosestTo(main2, splitPoint);
            var c5 = GetConnectorClosestTo(pipeCurrent, splitPoint);

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

        public static Connector GetConnectorClosestTo(Element e,
                                                      XYZ p)
        {
            ConnectorManager cm = GetConnectorManager(e);

            return null == cm
              ? null
              : GetConnectorClosestTo(cm.Connectors, p);
        }

        private static Connector GetConnectorClosestTo(ConnectorSet connectors,
                                                       XYZ p)
        {
            Connector targetConnector = null;
            double minDist = double.MaxValue;

            foreach (Connector c in connectors)
            {
                double d = c.Origin.DistanceTo(p);

                if (d < minDist)
                {
                    targetConnector = c;
                    minDist = d;
                }
            }
            return targetConnector;
        }

        private static ConnectorManager GetConnectorManager(Element e)
        {
            MEPCurve mc = e as MEPCurve;
            FamilyInstance fi = e as FamilyInstance;

            if (null == mc && null == fi)
            {
                throw new ArgumentException(
                  "Element is neither an MEP curve nor a fitting.");
            }

            return null == mc
              ? fi.MEPModel.ConnectorManager
              : mc.ConnectorManager;
        }

        public static bool se(MEPCurve mepCurveSplit1, MEPCurve mepCurveSplit2)
        {
            try
            {
                var locationCurve1 = mepCurveSplit1.GetCurve();
                var line1 = locationCurve1 as Line;

                var locationCurve2 = mepCurveSplit2.GetCurve();
                var line2 = locationCurve2 as Line;

                var p10 = line2.GetEndPoint(0);
                var p11 = line2.GetEndPoint(1);

                var inter1 = locationCurve1.Project(p10);
                var inter2 = locationCurve1.Project(p11);

                if (inter1 == null || inter2 == null)
                    return false;

                var d1 = inter1.XYZPoint.DistanceTo(p10);
                var d2 = inter2.XYZPoint.DistanceTo(p11);

                if (d1 < d2)
                {
                    var con = GetConnectorClosestTo(mepCurveSplit2, p10);
                    var elbow = Global.UIDoc.Document.Create.NewTakeoffFitting(con, mepCurveSplit1);
                }
                else
                {
                    var con = GetConnectorClosestTo(mepCurveSplit2, p11);
                    var elbow = Global.UIDoc.Document.Create.NewTakeoffFitting(con, mepCurveSplit1);
                }

                return true;
            }
            catch (System.Exception ex)
            {
                return false;
            }
        }

        public static FamilyInstance CreatTransitionFitting(MEPCurve mep1, MEPCurve mep2, bool checkDistance = false)
        {
            if (mep1 == null || mep2 == null)
                return null;

            var same = iss(mep1, mep2);
            if (same == true)
                return null;

            try
            {
                List<Connector> connectors = Common.GetConnectionNearest(mep1, mep2);
                if (connectors != null && connectors.Count == 2)
                {
                    var c0 = connectors[0];
                    var c1 = connectors[1];

                    if (checkDistance == true)
                    {
                        if (c0.Origin.DistanceTo(c1.Origin) > 1000 * Common.mmToFT)
                            return null;
                    }

                    if ((mep1 as Autodesk.Revit.DB.Mechanical.Duct != null || mep1 as Pipe != null) && (c0.IsConnected == true || c1.IsConnected == true))
                    {
                        return null;
                    }

                    var transition = Global.UIDoc.Document.Create.NewTransitionFitting(connectors[0], connectors[1]);
                    return transition;
                }
            }
            catch (System.Exception ex)
            {
            }
            return null;
        }

        public static bool iss(MEPCurve mep1, MEPCurve mep2)
        {
            //Check same type
            var shape1 = Common.GetShape(mep1);
            var shape2 = Common.GetShape(mep2);

            if (shape1 != shape2)
                return false;

            //Phai song song
            if (Common.IsParallel(mep1, mep2) == false)
            {
                return false;
            }

            //Check size: Size khac nhau moi tao

            if (mep1 is Autodesk.Revit.DB.Mechanical.Duct || mep1 is CableTray)
            {
                bool width = false;
                var paraW1 = mep1.LookupParameter("Width");
                var paraW2 = mep2.LookupParameter("Width");
                if (paraW1 != null && paraW2 != null)
                {
                    var d1 = paraW1.AsDouble();
                    var d2 = paraW2.AsDouble();
                    if (d1 == d2)
                        width = true;
                }

                bool height = false;
                var paraH1 = mep1.LookupParameter("Height");
                var paraH2 = mep2.LookupParameter("Height");
                if (paraH1 != null && paraH2 != null)
                {
                    var d1 = paraH1.AsDouble();
                    var d2 = paraH2.AsDouble();
                    if (d1 == d2)
                        height = true;
                }

                if (width == true && height == true)
                    return true;
            }
            else
            {
                var paraD1 = mep1.LookupParameter("Diameter");
                var paraD2 = mep2.LookupParameter("Diameter");
                if (paraD1 != null && paraD2 != null)
                {
                    var d1 = paraD1.AsDouble();
                    var d2 = paraD2.AsDouble();
                    if (d1 == d2)
                        return true;
                }

                paraD1 = mep1.LookupParameter("Diameter(Trade Size)");
                paraD2 = mep2.LookupParameter("Diameter(Trade Size)");
                if (paraD1 != null && paraD2 != null)
                {
                    var d1 = paraD1.AsDouble();
                    var d2 = paraD2.AsDouble();
                    if (d1 == d2)
                        return true;
                }
            }

            return false;
        }

        public static bool HandlerConnectAccessory(bool addNipple,
                                                   Pipe firstPipe,
                                                   Pipe primerPipe_1,
                                                   Pipe checkPrimerTemp1,
                                                   Pipe checkPrimerTemp2,
                                                   XYZ main1_p,
                                                   XYZ v_m_1,
                                                   XYZ main2_p,
                                                   XYZ v_m_2,
                                                   double diameter_10,
                                                   XYZ sub1Ints_pnt,
                                                   XYZ sub1_pnt0,
                                                   XYZ sub1_pnt1,
                                                   FamilyInstance tee,
                                                   out FamilyInstance reducer_1)
        {
            reducer_1 = null;
            if (primerPipe_1 != null)
            {
                if (checkPrimerTemp1 != null && primerPipe_1.Id == checkPrimerTemp1.Id)
                {
                    var line = Line.CreateUnbound(main1_p, v_m_1 * 10);

                    var p = line.Evaluate(diameter_10, false);
                    (primerPipe_1.Location as LocationCurve).Curve = Line.CreateBound(main1_p, p);

                    var futher = sub1Ints_pnt.DistanceTo(sub1_pnt0) > sub1Ints_pnt.DistanceTo(sub1_pnt1) ? sub1_pnt0 : sub1_pnt1;

                    SetLocationLine(firstPipe, futher, p);

                    reducer_1 = CreatTransitionFitting(firstPipe, primerPipe_1);
                }
                else if (checkPrimerTemp2 != null && primerPipe_1.Id == checkPrimerTemp2.Id)
                {
                    var line = Line.CreateUnbound(main2_p, v_m_2 * 10);

                    var p = line.Evaluate(diameter_10, false);
                    (primerPipe_1.Location as LocationCurve).Curve = Line.CreateBound(main2_p, p);

                    var futher = sub1Ints_pnt.DistanceTo(sub1_pnt0) > sub1Ints_pnt.DistanceTo(sub1_pnt1) ? sub1_pnt0 : sub1_pnt1;

                    SetLocationLine(firstPipe, futher, p);

                    reducer_1 = CreatTransitionFitting(firstPipe, primerPipe_1);
                }

                if (reducer_1 != null)
                    return true;
            }

            return false;
        }

        public static bool HandlerConnectAccessoryNipple(bool addNipple, Pipe primerPipe_1, FamilyInstance reducer_1, FamilyInstance tee)
        {
            try
            {
                if (addNipple)
                {
                    Tuple<Connector, Connector> cntsNipple_1 = PipeToPairConnector(primerPipe_1, reducer_1, tee);

                    Line lineNipple = Line.CreateBound(cntsNipple_1.Item1.Origin, cntsNipple_1.Item2.Origin);

                    XYZ center = MiddlePoint(cntsNipple_1.Item1.Origin, cntsNipple_1.Item2.Origin);

                    FamilyInstance nipple = Global.UIDoc.Document.Create.NewFamilyInstance(center, App._2LevelSmartForm.SelectedNippleFamily, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                    if (nipple != null)
                    {
                        var paraDia = primerPipe_1.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble();
                        nipple.LookupParameter("Nominal Diameter").Set(paraDia);
                        Global.UIDoc.Document.Delete(primerPipe_1.Id);
                    }

                    RotateLine(Global.UIDoc.Document, nipple, lineNipple);

                    List<Connector> temp_tee = GetConnectors(tee.MEPModel.ConnectorManager.Connectors, true);
                    temp_tee.Remove(temp_tee.OrderBy(item => item.Origin.Z).First());

                    List<Connector> temp1 = GetConnectors(nipple.MEPModel.ConnectorManager.Connectors, true);

                    Connector cnt_tee_1 = temp_tee[0];
                    Connector cnt_tee_2 = temp_tee[1];

                    Connector cnt_nipple_1 = temp1[0];
                    Connector cnt_nipple_2 = temp1[1];

                    List<Tuple<Tuple<Connector, Connector>, double>> distanceConts = new List<Tuple<Tuple<Connector, Connector>, double>>();
                    distanceConts.Add(new Tuple<Tuple<Connector, Connector>, double>(new Tuple<Connector, Connector>(cnt_tee_1, cnt_nipple_1), cnt_tee_1.Origin.DistanceTo(cnt_nipple_1.Origin)));
                    distanceConts.Add(new Tuple<Tuple<Connector, Connector>, double>(new Tuple<Connector, Connector>(cnt_tee_1, cnt_nipple_2), cnt_tee_1.Origin.DistanceTo(cnt_nipple_2.Origin)));
                    distanceConts.Add(new Tuple<Tuple<Connector, Connector>, double>(new Tuple<Connector, Connector>(cnt_tee_2, cnt_nipple_1), cnt_tee_2.Origin.DistanceTo(cnt_nipple_1.Origin)));
                    distanceConts.Add(new Tuple<Tuple<Connector, Connector>, double>(new Tuple<Connector, Connector>(cnt_tee_2, cnt_nipple_2), cnt_tee_2.Origin.DistanceTo(cnt_nipple_2.Origin)));

                    Tuple<Tuple<Connector, Connector>, double> minDistanceCnt = distanceConts.OrderBy(item => item.Item2).First();
                    XYZ translation = minDistanceCnt.Item1.Item1.Origin - minDistanceCnt.Item1.Item2.Origin;
                    ElementTransformUtils.MoveElement(Global.UIDoc.Document, nipple.Id, translation);

                    List<Connector> temp_reducer = GetConnectors(reducer_1.MEPModel.ConnectorManager.Connectors, true);

                    List<Connector> temp2 = GetConnectors(nipple.MEPModel.ConnectorManager.Connectors, true);

                    Connector cnt_reducer_1 = temp_reducer[0];
                    Connector cnt_reducer_2 = temp_reducer[1];

                    cnt_nipple_1 = temp2[0];
                    cnt_nipple_2 = temp2[1];

                    List<Tuple<Tuple<Connector, Connector>, double>> distanceContsReducer = new List<Tuple<Tuple<Connector, Connector>, double>>();
                    distanceContsReducer.Add(new Tuple<Tuple<Connector, Connector>, double>(new Tuple<Connector, Connector>(cnt_nipple_1, cnt_reducer_1), cnt_nipple_1.Origin.DistanceTo(cnt_reducer_1.Origin)));
                    distanceContsReducer.Add(new Tuple<Tuple<Connector, Connector>, double>(new Tuple<Connector, Connector>(cnt_nipple_1, cnt_reducer_2), cnt_nipple_1.Origin.DistanceTo(cnt_reducer_2.Origin)));
                    distanceContsReducer.Add(new Tuple<Tuple<Connector, Connector>, double>(new Tuple<Connector, Connector>(cnt_nipple_2, cnt_reducer_1), cnt_nipple_2.Origin.DistanceTo(cnt_reducer_1.Origin)));
                    distanceContsReducer.Add(new Tuple<Tuple<Connector, Connector>, double>(new Tuple<Connector, Connector>(cnt_nipple_2, cnt_reducer_2), cnt_nipple_2.Origin.DistanceTo(cnt_reducer_2.Origin)));

                    Tuple<Tuple<Connector, Connector>, double> minDistanceCntReducer = distanceContsReducer.OrderBy(item => item.Item2).First();
                    XYZ translation_2 = minDistanceCntReducer.Item1.Item1.Origin - minDistanceCntReducer.Item1.Item2.Origin;
                    ElementTransformUtils.MoveElement(Global.UIDoc.Document, reducer_1.Id, translation_2);

                    List<Connector> temp3 = GetConnectors(nipple.MEPModel.ConnectorManager.Connectors, true);
                    List<Connector> temp3_tee = GetConnectors(tee.MEPModel.ConnectorManager.Connectors, true);
                    List<Connector> temp3_reducer = GetConnectors(reducer_1.MEPModel.ConnectorManager.Connectors, true);
                    foreach (Connector cnt in temp3_tee)
                    {
                        if (Common.IsEqual(temp3[0].Origin, cnt.Origin))
                        {
                            temp3[0].ConnectTo(cnt);
                        }
                        else
                        {
                            temp3[1].ConnectTo(cnt);
                        }
                        break;
                    }

                    foreach (Connector cnt in temp3_reducer)
                    {
                        if (Common.IsEqual(temp3[0].Origin, cnt.Origin))
                        {
                            temp3[0].ConnectTo(cnt);
                        }
                        else
                        {
                            temp3[1].ConnectTo(cnt);
                        }
                        break;
                    }
                }
            }
            catch (Exception)
            { }
            return false;
        }

        public static Tuple<Connector, Connector> PipeToPairConnector(Pipe pipe, FamilyInstance reducer, FamilyInstance tee)
        {
            try
            {
                if (pipe == null || reducer == null || tee == null)
                    return new Tuple<Connector, Connector>(null, null);

                List<Connector> cntOfPipe = GetConnectors(pipe.ConnectorManager.Connectors, true);
                List<Connector> cntOfReducer = GetConnectors(reducer.MEPModel.ConnectorManager.Connectors, true);
                List<Connector> cntOfTee = GetConnectors(tee.MEPModel.ConnectorManager.Connectors, true);

                if (cntOfPipe.Count >= 2)
                {
                    Connector cntPipe_1 = cntOfPipe[0];
                    Connector cntPipe_2 = cntOfPipe[1];

                    Connector retCnt_1 = cntOfReducer.Where(item => Common.IsEqual(item.Origin, cntPipe_1.Origin) || Common.IsEqual(item.Origin, cntPipe_2.Origin)).FirstOrDefault();
                    Connector retCnt_2 = cntOfTee.Where(item => Common.IsEqual(item.Origin, cntPipe_1.Origin) || Common.IsEqual(item.Origin, cntPipe_2.Origin)).FirstOrDefault();

                    if (retCnt_1 != null && retCnt_2 != null)
                    {
                        return new Tuple<Connector, Connector>(retCnt_1, retCnt_2);
                    }
                }
            }
            catch (Exception)
            { }
            return new Tuple<Connector, Connector>(null, null);
        }

        private static List<Connector> GetConnectors(ConnectorSet connectorSet, bool filter = false)
        {
            try
            {
                List<Connector> retVal = new List<Connector>();
                foreach (Connector connector in connectorSet)
                {
                    if (connector.ConnectorType != ConnectorType.End && filter == true)
                        continue;
                    retVal.Add(connector);
                }
                return retVal;
            }
            catch (Exception)
            { }
            return new List<Connector>();
        }

        public static XYZ MiddlePoint(XYZ xYZ_1, XYZ xYZ_2) => new XYZ((xYZ_1.X + xYZ_2.X) * 0.5, (xYZ_1.Y + xYZ_2.Y) * 0.5, (xYZ_1.Z + xYZ_2.Z) * 0.5);

        public static void RotateLine(Document doc, FamilyInstance wye, Line axisLine)
        {
            var lst = Common.ToList(wye.MEPModel.ConnectorManager.Connectors);

            Connector connector2 = lst[0];
            Connector connector3 = lst[1];

            Line rotateLine = Line.CreateBound(connector2.Origin, connector3.Origin);

            if (Common.IsParallel(axisLine.Direction, rotateLine.Direction))
                return;

            XYZ vector = rotateLine.Direction.CrossProduct(axisLine.Direction);
            XYZ intersection = GetUnBoundIntersection(rotateLine, axisLine);

            double angle = rotateLine.Direction.AngleTo(axisLine.Direction);

            Line line = Line.CreateUnbound(intersection, vector);

            ElementTransformUtils.RotateElement(doc, wye.Id, line, angle);
            doc.Regenerate();
        }

        public static XYZ GetUnBoundIntersection(Line Line1, Line Line2)
        {
            if (Line1 != null && Line2 != null)
            {
                Curve ExtendedLine1 = Line.CreateUnbound(Line1.Origin, Line1.Direction);
                Curve ExtendedLine2 = Line.CreateUnbound(Line2.Origin, Line2.Direction);
                SetComparisonResult setComparisonResult = ExtendedLine1.Intersect(ExtendedLine2, out IntersectionResultArray resultArray);
                if (resultArray != null &&
                    resultArray.Size > 0)
                {
                    foreach (IntersectionResult result in resultArray)
                        if (result != null)
                            return result.XYZPoint;
                }
                else
                {
                    if (Line1.IsBound && Line2.IsBound)
                    {
                        if (Line1.GetEndPoint(0).IsAlmostEqualTo(Line2.GetEndPoint(0)) ||
                  Line1.GetEndPoint(0).IsAlmostEqualTo(Line2.GetEndPoint(1)))
                            return Line1.GetEndPoint(0);
                        else
                   if (Line1.GetEndPoint(1).IsAlmostEqualTo(Line2.GetEndPoint(0)) ||
                  Line1.GetEndPoint(1).IsAlmostEqualTo(Line2.GetEndPoint(1)))
                            return Line1.GetEndPoint(1);
                    }
                }
            }
            return null;
        }

        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

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

        private static List<PairPipes> PairPipe(List<Pipe> pipes)
        {
            //             if (pipes.Count < 2)
            //                 return null;

            double temp = 100;

            List<PairPipes> pairs = new List<PairPipes>();

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
                            var newPair = new PairPipes(pipe, pipe2);
                            pairs.Add(newPair);
                        }
                    }
                }
            }

            List<ElementId> allAddeds1 = (from PairPipes pair in pairs
                                          where pair._Pipe1 != null
                                          select pair._Pipe1.Id).ToList();

            List<ElementId> allAddeds2 = (from PairPipes pair in pairs
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

                var newPair = new PairPipes(pipe, null);
                pairs.Add(newPair);
            }

            return pairs;
        }

        private static bool ce(PairPipes pair, Pipe pipe1, Pipe pipe2)
        {
            if (pair._Pipe1.Id == pipe1.Id && pair._Pipe2.Id == pipe2.Id)
                return true;
            if (pair._Pipe1.Id == pipe2.Id && pair._Pipe2.Id == pipe1.Id)
                return true;

            return false;
        }

        private static bool ps(Pipe pipe1, Pipe pipe2, double diameter, PipeType typePipe)
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

                foreach (ElementId pId in m_mainPipes)
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

                        var fitting = CreateTeeFitting(main as Pipe, vertical as Pipe, inter_main1, out main2);
                        if (main2 != null && fitting != null)
                        {
                            m_mainPipes.Add(main2.Id);
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

                                SetLocationLine(pipe1, futher, p);

                                CreateFittingForMEPUtils.CreatTransitionFitting(pipe1, pipe_moi_1);
                            }
                            else if (mep2 != null && pipe_moi_1.Id == mep2.Id)
                            {
                                var line = Line.CreateUnbound(main2_p, v_m_2 * 10);

                                var p = line.Evaluate(d10, false);
                                (pipe_moi_1.Location as LocationCurve).Curve = Line.CreateBound(main2_p, p);

                                var futher = inters_sub1.DistanceTo(p0_sub1) > inters_sub1.DistanceTo(p1_sub1) ? p0_sub1 : p1_sub1;

                                SetLocationLine(pipe1, futher, p);

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

                                SetLocationLine(pipe2, futher, p);

                                CreateFittingForMEPUtils.CreatTransitionFitting(pipe2, pipe_moi_2);
                            }
                            else if (mep2 != null && pipe_moi_2.Id == mep2.Id)
                            {
                                var line = Line.CreateUnbound(main2_p, v_m_2 * 10);

                                var p = line.Evaluate(d10, false);
                                (pipe_moi_2.Location as LocationCurve).Curve = Line.CreateBound(main2_p, p);

                                var futher = inters_sub2.DistanceTo(p0_sub2) > inters_sub2.DistanceTo(p1_sub2) ? p0_sub2 : p1_sub2;

                                SetLocationLine(pipe2, futher, p);

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

        private static void SetLocationLine(Pipe pipe, XYZ pOn, XYZ pOther)
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

        private static Pipe ee(ConnectorSet cs)
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

        private static bool g(double d1, double d2)
        {
            if (Math.Abs(d1 - d2) < 0.001)
                return true;

            return false;
        }

        private static PreferredJunctionType gr(Pipe pipe)
        {
            var pipeType = pipe.PipeType as PipeType;

            return pipeType.RoutingPreferenceManager.PreferredJunctionType;
        }

        private static bool dv(Pipe main, Pipe pipeSub, bool twoPipes, out XYZ inters_main, out XYZ p0_sub, out XYZ p1_sub, out XYZ inters_sub)
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

    public class PairPipes
    {
        public Pipe _Pipe1 = null;
        public Pipe _Pipe2 = null;

        public PairPipes(Pipe p1, Pipe p2)
        {
            _Pipe1 = p1;
            _Pipe2 = p2;
        }
    }
}