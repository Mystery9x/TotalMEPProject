using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using TotalMEPProject.Services;
using TotalMEPProject.Ultis;

namespace TotalMEPProject.Commands.FireFighting
{
    [Transaction(TransactionMode.Manual)]
    public class CmdTwoLevelSmart : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Global.UIApp = commandData.Application;
            Global.RVTApp = commandData.Application.Application;
            Global.UIDoc = commandData.Application.ActiveUIDocument;
            Global.AppCreation = commandData.Application.Application.Create;

            //Show dialog
            if (App.Show2LevelSmartForm() == false)
                return Result.Cancelled;

            return Result.Succeeded;
        }

        /// <summary>
        /// Select multies pipe
        /// </summary>
        /// <returns></returns>
        public static List<Pipe> SelectPipes()
        {
            List<Pipe> retVal = new List<Pipe>();
            try
            {
                var pickedObjs = Global.UIDoc.Selection.PickObjects(ObjectType.Element, new MEPCurveFilter/*PipeFilter*/(), "SELECT PIPES : ");

                foreach (Reference pickedObj in pickedObjs)
                {
                    var pipe = Global.UIDoc.Document.GetElement(pickedObj) as Pipe;

                    if (pipe != null)
                        retVal.Add(pipe);
                }
            }
            catch (System.Exception ex)
            {
            }
            return retVal;
        }

        /// <summary>
        /// Pair pipe
        /// </summary>
        /// <param name="pipes"></param>
        /// <returns></returns>
        private static List<PairPipes> PairingPipes(List<Pipe> pipes)
        {
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

                        var find = pairs.Find(item => IsEqualPipe(item, pipe, pipe2));
                        if (find == null)
                        {
                            var newPair = new PairPipes(pipe, pipe2);
                            pairs.Add(newPair);
                        }
                    }
                }
            }

            List<ElementId> allAddeds1 = (from PairPipes pair in pairs
                                          where pair.Pipe1 != null
                                          select pair.Pipe1.Id).ToList();

            List<ElementId> allAddeds2 = (from PairPipes pair in pairs
                                          where pair.Pipe2 != null
                                          select pair.Pipe2.Id).ToList();

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

        private static List<DuoBranchPipe> PairingPipesPlus(List<Pipe> pipes)
        {
            List<InforPie> inforPipes = new List<InforPie>();
            pipes.ForEach(item => inforPipes.Add(new InforPie(item)));

            List<DuoBranchPipe> pairs = new List<DuoBranchPipe>();

            foreach (InforPie inforPipe1 in inforPipes)
            {
                var curve_sub = inforPipe1.CurveSourcePipe;

                var expand = inforPipe1.CurveSourcePipe_Extend;

                foreach (InforPie inforPipe2 in inforPipes)
                {
                    if (inforPipe1.SourcePipe.Id == inforPipe2.SourcePipe.Id)
                        continue;

                    var curve_sub2 = inforPipe2.CurveSourcePipe;
                    IntersectionResultArray arr = new IntersectionResultArray();
                    var result = curve_sub2.Intersect(expand, out arr);

                    if (result == SetComparisonResult.Equal)
                    {
                        //Dong tam
                        var find = pairs.Find(item => IsEqualPipe(item, inforPipe1, inforPipe2));
                        if (find == null)
                        {
                            var newPair = new DuoBranchPipe(inforPipe1, inforPipe2);
                            pairs.Add(newPair);
                        }
                    }
                }
            }

            List<ElementId> allAddeds1 = (from DuoBranchPipe pair in pairs
                                          where pair.FirstPipe.SourcePipe != null
                                          select pair.FirstPipe.SourcePipe.Id).ToList();

            List<ElementId> allAddeds2 = (from DuoBranchPipe pair in pairs
                                          where pair.SecondPipe.SourcePipe != null
                                          select pair.SecondPipe.SourcePipe.Id).ToList();

            List<ElementId> rest = (from Pipe pipe in pipes
                                    where allAddeds1.Contains(pipe.Id) == false && allAddeds2.Contains(pipe.Id) == false
                                    select pipe.Id).ToList();

            foreach (ElementId id in rest)
            {
                var pipe = Global.UIDoc.Document.GetElement(id) as Pipe;
                if (pipe == null)
                    continue;

                var newPair = new DuoBranchPipe(new InforPie(pipe), null);
                pairs.Add(newPair);
            }

            return pairs;
        }

        /// <summary>
        /// Is Equal Pipe
        /// </summary>
        /// <param name="pair"></param>
        /// <param name="pipe1"></param>
        /// <param name="pipe2"></param>
        /// <returns></returns>
        private static bool IsEqualPipe(PairPipes pair, Pipe pipe1, Pipe pipe2)
        {
            if (pair.Pipe1.Id == pipe1.Id && pair.Pipe2.Id == pipe2.Id)
                return true;
            if (pair.Pipe1.Id == pipe2.Id && pair.Pipe2.Id == pipe1.Id)
                return true;

            return false;
        }

        private static bool IsEqualPipe(DuoBranchPipe pair, InforPie pipe1, InforPie pipe2)
        {
            if (pair.FirstPipe.SourcePipe.Id == pipe1.SourcePipe.Id && pair.SecondPipe.SourcePipe.Id == pipe2.SourcePipe.Id)
                return true;
            if (pair.FirstPipe.SourcePipe.Id == pipe2.SourcePipe.Id && pair.SecondPipe.SourcePipe.Id == pipe1.SourcePipe.Id)
                return true;

            return false;
        }

        /// <summary>
        /// Process
        /// </summary>
        /// <returns></returns>
        public static Result Process()
        {
            try
            {
                if (App.m_2LevelSmartForm != null && App.m_2LevelSmartForm.IsDisposed == false)
                {
                    App.m_2LevelSmartForm.Hide();
                }

                // Select main pipes
                List<Pipe> mainPipes = SelectPipes();
                if (mainPipes == null || mainPipes.Count <= 0)
                    return Result.Cancelled;

                List<InforPie> inforMainPipes = new List<InforPie>();
                mainPipes.ForEach(item => inforMainPipes.Add(new InforPie(item)));

                // Select branch pipes
                List<Pipe> branchPipes = SelectPipes();
                if (branchPipes == null || branchPipes.Count <= 0)
                    return Result.Cancelled;

                // Filter branch pipes
                branchPipes = branchPipes.Where(item => (mainPipes.Find(item_1 => item_1.Id == item.Id) == null)).Select(item => item).ToList();

                // Pair pipe
                List<DuoBranchPipe> pairPies = PairingPipesPlus(branchPipes);
                if (pairPies == null || pairPies.Count() <= 0)
                    return Result.Cancelled;

                TransactionGroup reTransGrp = new TransactionGroup(Global.UIDoc.Document, "TWO_LEVEL_SMART_COMMAND");
                try
                {
                    reTransGrp.Start();

                    if (GetPreferredJunctionType(mainPipes[0]) == PreferredJunctionType.Tap)
                    {
                        ProcessWithTap(inforMainPipes, pairPies, App.m_2LevelSmartForm.DialogResultData);
                    }
                    else if (GetPreferredJunctionType(mainPipes[0]) == PreferredJunctionType.Tee)
                    {
                    }
                    reTransGrp.Assimilate();
                }
                catch (Exception)
                {
                    reTransGrp.RollBack();
                }
            }
            catch (Exception)
            { }
            finally
            {
                if (App.m_2LevelSmartForm != null && App.m_2LevelSmartForm.IsDisposed == false)
                {
                    App.m_2LevelSmartForm.Show(App.hWndRevit);
                }
                DisplayService.SetFocus(new HandleRef(null, App.m_2LevelSmartForm.Handle));
            }
            return Result.Cancelled;
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="pipe"></param>
        /// <returns></returns>
        private static PreferredJunctionType GetPreferredJunctionType(Pipe pipe)
        {
            var pipeType = pipe.PipeType as PipeType;

            return pipeType.RoutingPreferenceManager.PreferredJunctionType;
        }

        private static bool ProcessWithTap(List<InforPie> inforMainPipes, List<DuoBranchPipe> pairPies, TwoLevelSmartDialogData dialogResultData)
        {
            try
            {
                List<InforPie> process_InforMains = new List<InforPie>(inforMainPipes);
                List<DuoBranchPipe> process_duoBranchPipes = new List<DuoBranchPipe>(pairPies);

                List<SourceMainPipe> sourceMainPipes = new List<SourceMainPipe>();

                foreach (var inforMain in process_InforMains)
                {
                    var flatten_mainCurve = inforMain.CurveSourcePipe_Flatten;
                    var flatten_mainCurve_Extend = inforMain.CuvreSourcePipe_FlattenExtend;
                    SourceMainPipe sourceMain = new SourceMainPipe(inforMain);

                    foreach (var duoBranchs in process_duoBranchPipes)
                    {
                        if (RealityIntersect(flatten_mainCurve, duoBranchs.FlattenValidCurve, out XYZ flattenIntPnt1))
                        {
                            DuoBranchPipe duoBranchPipe = new DuoBranchPipe(duoBranchs.FirstPipe, duoBranchs.SecondPipe)
                            {
                                FlattenIntersectPoint = flattenIntPnt1
                            };
                            sourceMain.Branches.Add(duoBranchPipe);
                        }

                        if (RealityIntersect(flatten_mainCurve_Extend, duoBranchs.FlattenValidCurve, out XYZ flattenIntPnt2))
                        {
                            DuoBranchPipe duoBranchPipe = new DuoBranchPipe(duoBranchs.FirstPipe, duoBranchs.SecondPipe)
                            {
                                FlattenIntersectPoint = flattenIntPnt2
                            };
                            sourceMain.Branches_Special.Add(duoBranchPipe);
                        }
                    }

                    sourceMainPipes.Add(sourceMain);
                }
            }
            catch (Exception)
            { }
            return false;
        }

        private static bool RealityIntersect(Line mainLine, Line checkLine, out XYZ intersectPnt)
        {
            intersectPnt = null;
            try
            {
                IntersectionResultArray intsRetArr = new IntersectionResultArray();
                var intsRet = mainLine.Intersect(checkLine, out intsRetArr);
                if (intsRet == SetComparisonResult.Overlap)
                {
                    intersectPnt = intsRetArr.get_Item(0).XYZPoint;
                    return true;
                }
            }
            catch (Exception)
            { }
            return false;
        }
    }

    public class InforPie
    {
        public Pipe SourcePipe { get; set; }

        public Connector SourcePipe_Cont_1
        {
            get
            {
                if (SourcePipe != null && SourcePipe.IsValidObject && CurveSourcePipe != null)
                    return Common.GetConnectorClosestTo(SourcePipe, CurveSourcePipe.GetEndPoint(0));
                return null;
            }
        }

        public Connector SourcePipe_Cont_2
        {
            get
            {
                if (SourcePipe != null && SourcePipe.IsValidObject && CurveSourcePipe != null)
                    return Common.GetConnectorClosestTo(SourcePipe, CurveSourcePipe.GetEndPoint(1));
                return null;
            }
        }

        public Line CurveSourcePipe
        {
            get
            {
                if (SourcePipe != null && SourcePipe.IsValidObject)
                {
                    return SourcePipe.GetCurve() as Line;
                }
                return null;
            }
        }

        public Line CurveSourcePipe_Unbound
        {
            get
            {
                if (SourcePipe != null && SourcePipe.IsValidObject && CurveSourcePipe != null)
                {
                    return Line.CreateUnbound(CurveSourcePipe.GetEndPoint(0), CurveSourcePipe.Direction);
                }
                return null;
            }
        }

        public Line CurveSourcePipe_Extend
        {
            get
            {
                if (SourcePipe != null && SourcePipe.IsValidObject && SourcePipe_Cont_1 != null && SourcePipe_Cont_2 != null)
                {
                    if (!SourcePipe_Cont_1.IsConnected && SourcePipe_Cont_2.IsConnected)
                    {
                        XYZ newStartPoint = CurveSourcePipe.Evaluate(CurveSourcePipe.GetEndParameter(0) - 1000 * Common.mmToFT, false);
                        XYZ newEndPoint = CurveSourcePipe.GetEndPoint(1);
                        return Line.CreateBound(newStartPoint, newEndPoint);
                    }
                    else if (SourcePipe_Cont_1.IsConnected && !SourcePipe_Cont_2.IsConnected)
                    {
                        XYZ newStartPoint = CurveSourcePipe.GetEndPoint(0);
                        XYZ newEndPoint = CurveSourcePipe.Evaluate(CurveSourcePipe.GetEndParameter(1) + 1000 * Common.mmToFT, false);
                        return Line.CreateBound(newStartPoint, newEndPoint);
                    }
                    else if (SourcePipe_Cont_1.IsConnected && SourcePipe_Cont_2.IsConnected)
                    {
                        return CurveSourcePipe;
                    }
                    else
                    {
                        XYZ newStartPoint = CurveSourcePipe.Evaluate(CurveSourcePipe.GetEndParameter(0) - 1000 * Common.mmToFT, false);
                        XYZ newEndPoint = CurveSourcePipe.Evaluate(CurveSourcePipe.GetEndParameter(1) + 600 * Common.mmToFT, false);
                        return Line.CreateBound(newStartPoint, newEndPoint);
                    }
                }
                return null;
            }
        }

        public Line CurveSourcePipe_Flatten
        {
            get
            {
                if (SourcePipe != null && SourcePipe.IsValidObject)
                {
                    return SourcePipe.GetFlattenCurve() as Line;
                }
                return null;
            }
        }

        public Line CuvreSourcePipe_FlattenExtend
        {
            get
            {
                if (SourcePipe != null && SourcePipe.IsValidObject && SourcePipe_Cont_1 != null && SourcePipe_Cont_2 != null)
                {
                    if (SourcePipe_Cont_1.IsConnected && !SourcePipe_Cont_2.IsConnected)
                    {
                        XYZ newStartPoint = CurveSourcePipe_Flatten.Evaluate(CurveSourcePipe_Flatten.GetEndParameter(0) - 600 * Common.mmToFT, false);
                        XYZ newEndPoint = CurveSourcePipe_Flatten.GetEndPoint(1);
                        return Line.CreateBound(newStartPoint, newEndPoint);
                    }
                    else if (!SourcePipe_Cont_1.IsConnected && SourcePipe_Cont_2.IsConnected)
                    {
                        XYZ newStartPoint = CurveSourcePipe_Flatten.GetEndPoint(0);
                        XYZ newEndPoint = CurveSourcePipe_Flatten.Evaluate(CurveSourcePipe_Flatten.GetEndParameter(1) + 600 * Common.mmToFT, false);
                        return Line.CreateBound(newStartPoint, newEndPoint);
                    }
                    else if (SourcePipe_Cont_1.IsConnected && SourcePipe_Cont_2.IsConnected)
                    {
                        return CurveSourcePipe_Flatten;
                    }
                    else
                    {
                        XYZ newStartPoint = CurveSourcePipe_Flatten.Evaluate(CurveSourcePipe_Flatten.GetEndParameter(0) - 600 * Common.mmToFT, false);
                        XYZ newEndPoint = CurveSourcePipe_Flatten.Evaluate(CurveSourcePipe_Flatten.GetEndParameter(1) + 600 * Common.mmToFT, false);
                        return Line.CreateBound(newStartPoint, newEndPoint);
                    }
                }
                return null;
            }
        }

        public InforPie(Pipe sourcePipe)
        {
            SourcePipe = sourcePipe;
        }
    }

    public class DuoBranchPipe
    {
        public InforPie FirstPipe { get; set; }
        public InforPie SecondPipe { get; set; }
        public XYZ FlattenIntersectPoint { get; set; }

        public DuoBranchPipe(InforPie firstPipe, InforPie secondPipe)
        {
            FirstPipe = firstPipe;
            SecondPipe = secondPipe;
        }

        public Line FlattenValidCurve
        {
            get
            {
                if (FirstPipe != null && SecondPipe != null && FirstPipe.SourcePipe != null && SecondPipe.SourcePipe != null)
                {
                    XYZ f_p1 = FirstPipe.SourcePipe_Cont_1.Origin;
                    XYZ f_p2 = FirstPipe.SourcePipe_Cont_2.Origin;
                    XYZ s_p1 = SecondPipe.SourcePipe_Cont_1.Origin;
                    XYZ s_p2 = SecondPipe.SourcePipe_Cont_2.Origin;
                    List<XYZ> xYZs = new List<XYZ>() { f_p1, f_p2, s_p1, s_p2 };

                    XYZ point1 = xYZs[0];
                    XYZ point2 = xYZs[1];
                    double maxDistance = point1.DistanceTo(point2);
                    for (int i = 0; i < xYZs.Count; i++)
                    {
                        for (int j = i + 1; j < xYZs.Count; j++)
                        {
                            double distance = xYZs[i].DistanceTo(xYZs[j]);
                            if (distance > maxDistance)
                            {
                                maxDistance = distance;
                                point1 = xYZs[i];
                                point2 = xYZs[j];
                            }
                        }
                    }

                    return Line.CreateBound(point1.FlattenPoint(), point2.FlattenPoint());
                }
                else if (FirstPipe != null && SecondPipe == null && FirstPipe.SourcePipe != null)
                {
                    return FirstPipe.CurveSourcePipe_Extend;
                }
                return null;
            }
        }
    }

    public class PointComparer : IComparer<XYZ>
    {
        private XYZ axis;
        private XYZ vector;
        private XYZ furthestPoint1;

        public PointComparer(XYZ furthestPoint1, XYZ furthestPoint2)
        {
            this.furthestPoint1 = furthestPoint1;
            vector = furthestPoint2 - furthestPoint1;
            axis = vector.CrossProduct(XYZ.BasisZ);
        }

        public int Compare(XYZ p1, XYZ p2)
        {
            XYZ v1 = p1 - furthestPoint1;
            XYZ v2 = p2 - furthestPoint1;
            double angle1 = Math.Atan2(v1.DotProduct(axis), v1.DotProduct(vector));
            double angle2 = Math.Atan2(v2.DotProduct(axis), v2.DotProduct(vector));
            return angle1.CompareTo(angle2);
        }
    }

    public class SourceMainPipe
    {
        public InforPie MainPipe { get; set; }
        public List<DuoBranchPipe> Branches { get; set; }
        public List<DuoBranchPipe> Branches_Special { get; set; }

        public SourceMainPipe(InforPie mainPipe)
        {
            MainPipe = mainPipe;
            Branches = new List<DuoBranchPipe>();
        }
    }

    public static class PipeExtension
    {
        public static XYZ FlattenPoint(this XYZ point3d, double z = 0)
        {
            return new XYZ(point3d.X, point3d.Y, z);
        }

        public static List<Connector> GetConectors(this ConnectorSet connectors)
        {
            List<Connector> connects = new List<Connector>();
            foreach (Connector connector in connectors)
            {
                connects.Add(connector);
            }
            return connects;
        }

        public static Curve GetCurve(this Element element)
        {
            Debug.Assert(null != element.Location,
              "Expected an element with a valid Location");

            LocationCurve locationCurve = element.Location as LocationCurve;

            Debug.Assert(null != locationCurve,
              "Expected an element with a valid LocationCurve");

            return locationCurve.Curve;
        }

        public static Line GetFlattenCurve(this Element element)
        {
            Curve realCurve = element.GetCurve();
            Line retLine = Line.CreateBound(new XYZ(realCurve.GetEndPoint(0).X, realCurve.GetEndPoint(0).Y, 0), new XYZ(realCurve.GetEndPoint(1).X, realCurve.GetEndPoint(1).Y, 0));
            return retLine;
        }

        public static Line GetExpandCurve(this Element element, double expandVal = 0)
        {
            Curve realCurve = element.GetCurve();
            Line lineFromCurve = realCurve as Line;

            XYZ newStartPoint = lineFromCurve.Evaluate(lineFromCurve.GetEndParameter(0) - expandVal, false);
            XYZ newEndPoint = lineFromCurve.Evaluate(lineFromCurve.GetEndParameter(1) - expandVal, false);

            return Line.CreateBound(newStartPoint, newEndPoint);
        }

        public static Line GetExpandFlattenCurve(this Element element, double expandVal = 0)
        {
            Curve realCurve = element.GetFlattenCurve();
            Line lineFromCurve = realCurve as Line;

            XYZ newStartPoint = lineFromCurve.Evaluate(lineFromCurve.GetEndParameter(0) - expandVal, false);
            XYZ newEndPoint = lineFromCurve.Evaluate(lineFromCurve.GetEndParameter(1) + expandVal, false);

            return Line.CreateBound(newStartPoint, newEndPoint);
        }
    }
}