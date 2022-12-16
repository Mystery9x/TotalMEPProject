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
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
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
                SourcePipesData sourcePipesData = new SourcePipesData(m_mainPipes, pipe1, pipe2, pipeType, diameter, App._2LevelSmartForm.OptionAddNipple, App._2LevelSmartForm.SelectedNippleFamily, App._2LevelSmartForm.OptionAddElbowConnection);
                //HandlerWithPipes(pipe1, pipe2, diameter, pipeType);
            }

            t.Commit();

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

    public class SourcePipesData
    {
        #region Variable

        private Pipe m_firstPipe = null;
        private Pipe m_secondPipe = null;
        private Pipe m_mainPipe = null;
        private PipeType m_pipeTypeProcess = null;
        private double m_diameter = double.MinValue;
        private bool m_addNipple = false;
        private FamilySymbol m_nippleFamily = null;
        private bool m_addElbowConnection = false;
        private List<ElementId> m_mainPipes = new List<ElementId>();
        private Pipe m_verticalPipe = null;

        private XYZ m_intsMain1_pnt = null;
        private XYZ m_sub1_pnt0 = null;
        private XYZ m_sub1_pnt1 = null;
        private XYZ m_sub1Ints_pnt = null;

        private XYZ m_intsMain2_pnt = null;
        private XYZ m_sub2_pnt0 = null;
        private XYZ m_sub2_pnt1 = null;
        private XYZ m_sub2Ints_pnt = null;

        private FamilyInstance m_bottomTee = null;
        private FamilyInstance m_bottomElbow = null;

        private Pipe m_primerPipe_1 = null;
        private Pipe m_primerPipe_2 = null;

        private FamilyInstance m_topTee = null;

        private FamilyInstance m_topReducer1 = null;
        private FamilyInstance m_topReducer2 = null;

        #endregion Variable

        #region Properties

        public Pipe FirstPipe { get => m_firstPipe; set => m_firstPipe = value; }
        public Pipe SecondPipe { get => m_secondPipe; set => m_secondPipe = value; }
        public Pipe MainPipe { get => m_mainPipe; set => m_mainPipe = value; }
        public PipeType PipeTypeProcess { get => m_pipeTypeProcess; set => m_pipeTypeProcess = value; }
        public List<ElementId> MainPipes { get => m_mainPipes; set => m_mainPipes = value; }
        public double Diameter { get => m_diameter; set => m_diameter = value; }
        public bool FlagAddNipple { get => m_addNipple; set => m_addNipple = value; }
        public FamilySymbol NippleFamily { get => m_nippleFamily; set => m_nippleFamily = value; }
        public bool FlagAddElbowConnection { get => m_addElbowConnection; set => m_addElbowConnection = value; }
        public Pipe VerticalPipe { get => m_verticalPipe; set => m_verticalPipe = value; }

        #endregion Properties

        #region Constructor

        public SourcePipesData(List<ElementId> mainPipes, Pipe firstPipe, Pipe secondPipe, PipeType pipeType, double pipeDiameter, bool flagAddNipple, FamilySymbol nippleFamily, bool flagAddElbowCnt)
        {
            if (mainPipes != null)
                MainPipes = mainPipes;

            FirstPipe = firstPipe;
            SecondPipe = secondPipe;
            PipeTypeProcess = pipeType;
            Diameter = pipeDiameter;
            FlagAddNipple = flagAddNipple;
            NippleFamily = nippleFamily;
            FlagAddElbowConnection = flagAddElbowCnt;
            Initialize();
        }

        #endregion Constructor

        #region Method

        private void Initialize()
        {
            using (SubTransaction reSubTrans = new SubTransaction(Global.UIDoc.Document))
            {
                try
                {
                    reSubTrans.Start();

                    if (BeforeProcess() == false)
                    {
                        reSubTrans.RollBack();
                        return;
                    }
                    if (HandleVerticalPipe() == false)
                    {
                        reSubTrans.RollBack();
                        return;
                    }

                    if (FlagAddElbowConnection)
                    {
                        if (HandleFlagElbowConnection() == false)
                        {
                            reSubTrans.RollBack();
                            return;
                        }
                    }
                    else
                    {
                        if (HandlerBottomTee() == false)
                        {
                            reSubTrans.RollBack();
                            return;
                        }
                    }

                    if (HandlerTopTee() == false)
                    {
                        reSubTrans.RollBack();
                        return;
                    }

                    reSubTrans.Commit();
                }
                catch (Exception)
                {
                    reSubTrans.RollBack();
                }
            }
        }

        public bool BeforeProcess()
        {
            try
            {
                // Check has two pipe valid
                bool flagTwoPipe = SecondPipe != null ? true : false;

                // Get main pipe
                MainPipe = null;
                if (MainPipes == null || MainPipes.Count == 0)
                    return false;

                foreach (ElementId eId in m_mainPipes)
                {
                    Pipe pipeLoop = Global.UIDoc.Document.GetElement(eId) as Pipe;
                    if (pipeLoop == null)
                        continue;

                    if (HandlerDivide(pipeLoop, FirstPipe, flagTwoPipe, out m_intsMain1_pnt, out m_sub1_pnt0, out m_sub1_pnt1, out m_sub1Ints_pnt) == true)
                    {
                        MainPipe = pipeLoop;
                        break;
                    }
                }

                // Set branch pipe perpendiculer main pipe
                if (HandleBranchPipePerpendiculerMainPipe() == false)
                    return false;

                // If only 1 pipe
                if (SecondPipe == null)
                {
                    (FirstPipe.Location as LocationCurve).Curve = Line.CreateBound(m_sub1_pnt0, m_sub1Ints_pnt);

                    SecondPipe = Common.Clone(FirstPipe) as Pipe;
                    (SecondPipe.Location as LocationCurve).Curve = Line.CreateBound(m_sub1Ints_pnt, m_sub1_pnt1);

                    // Set point value again
                    m_sub1_pnt1 = m_sub1Ints_pnt;
                }

                if (HandlerDivide(MainPipe, SecondPipe, true, out m_intsMain2_pnt, out m_sub2_pnt0, out m_sub2_pnt1, out m_sub2Ints_pnt) == false)
                {
                    return false;
                }

                if (m_intsMain1_pnt.DistanceTo(m_intsMain1_pnt) > 0.001)
                {
                    return false;
                }

                if (m_sub1Ints_pnt.DistanceTo(m_sub2Ints_pnt) > 0.001)
                {
                    return false;
                }

                return true;
            }
            catch (Exception)
            { }
            return false;
        }

        public bool HandleVerticalPipe()
        {
            try
            {
                VerticalPipe = Common.Clone(FirstPipe) as Pipe;
                VerticalPipe.PipeType = PipeTypeProcess;
                VerticalPipe.LookupParameter("Diameter").Set(Diameter);
                (VerticalPipe.Location as LocationCurve).Curve = Line.CreateBound(m_intsMain1_pnt, m_sub1Ints_pnt);

                return true;
            }
            catch (Exception)
            { }
            return false;
        }

        public bool HandleBranchPipePerpendiculerMainPipe()
        {
            try
            {
                if (FirstPipe == null || MainPipe == null)
                    return false;

                return true;
            }
            catch (Exception)
            { }
            return false;
        }

        public bool HandleFlagElbowConnection()
        {
            try
            {
                if (MainPipe as Pipe != null && VerticalPipe as Pipe != null && m_intsMain1_pnt != null)
                {
                    Pipe main2 = null;

                    double newDiameter = GetParameterValueByName(FirstPipe, "Diameter")
                                    != null ? (double)GetParameterValueByName(FirstPipe, "Diameter") : double.MinValue;

                    if (newDiameter == double.MinValue)
                    {
                        return false;
                    }

                    Diameter = newDiameter;

                    VerticalPipe.LookupParameter("Diameter").Set(newDiameter);
                    (VerticalPipe.Location as LocationCurve).Curve = Line.CreateBound(m_intsMain1_pnt, m_sub1Ints_pnt);

                    m_bottomElbow = CreateElbowFitting(MainPipe as Pipe, VerticalPipe as Pipe, m_intsMain1_pnt, out main2);
                    if (main2.IsValidObject == true && m_bottomElbow != null)
                    {
                        MainPipes.Add(main2.Id);
                    }
                    else if ((main2.IsValidObject != true && m_bottomElbow != null))
                    {
                    }
                    else
                    {
                        return false;
                    }
                    return true;
                }
            }
            catch (Exception)
            { }
            return false;
        }

        public bool HandlerBottomTee()
        {
            try
            {
                if (GetPreferredJunctionType(MainPipe) == PreferredJunctionType.Tee)
                {
                    if (MainPipe as Pipe != null && VerticalPipe as Pipe != null && m_intsMain1_pnt != null)
                    {
                        Pipe main2 = null;

                        m_bottomTee = CreateTeeFitting(MainPipe as Pipe, VerticalPipe as Pipe, m_intsMain1_pnt, out main2);
                        if (main2 != null && m_bottomTee != null)
                        {
                            MainPipes.Add(main2.Id);
                        }
                        else
                        {
                            return false;
                        }
                        return true;
                    }
                }
                else
                {
                    if (se(MainPipe, VerticalPipe) == false)
                    {
                        return false;
                    }
                }
            }
            catch (Exception)
            { }
            return false;
        }

        public bool HandlerTopTee()
        {
            try
            {
                // Process with sub pipes

                Connector cntBottom_topTee = GetConnectorClosestTo(VerticalPipe, m_sub1Ints_pnt);

                //double diameter_primer = 15 * Common.mmToFT;

                double diameter_10 = 10 * Common.mmToFT;

                // If Diameter vertical pipe = Diameter main pipe
                if (g(Diameter, FirstPipe.Diameter) && g(Diameter, SecondPipe.Diameter))
                {
                    //Connect to sub
                    (FirstPipe.Location as LocationCurve).Curve = Line.CreateBound(m_sub1_pnt0, m_sub1_pnt1);
                    (SecondPipe.Location as LocationCurve).Curve = Line.CreateBound(m_sub2_pnt0, m_sub2_pnt1);

                    //Connect
                    Connector c3 = GetConnectorClosestTo(FirstPipe, m_sub1Ints_pnt);
                    Connector c4 = GetConnectorClosestTo(SecondPipe, m_sub2Ints_pnt);
                    m_topTee = CreatTee(c3, c4, cntBottom_topTee);
                    if (m_topTee == null)
                    {
                        return false;
                    }

                    return true;
                }
                // If Diameter vertical pipe < Diameter main pipe
                else
                {
                    Connector c3 = null;
                    // Create primer pipe 1
                    m_primerPipe_1 = null;

                    XYZ tempVector = null;

                    if (g(Diameter, FirstPipe.Diameter) == false)
                    {
                        // Create primer pipe 1
                        m_primerPipe_1 = Common.Clone(VerticalPipe) as Pipe;
                        Line newLocationCurve = Line.CreateBound(m_sub1_pnt0, m_sub1_pnt1);

                        XYZ p1 = m_sub1_pnt0;

                        if (m_sub1_pnt0.DistanceTo(m_sub1Ints_pnt) < 0.01)
                        {
                            p1 = m_sub1_pnt1;
                        }

                        (m_primerPipe_1.Location as LocationCurve).Curve = Line.CreateBound(m_sub1Ints_pnt, p1);

                        c3 = Common.GetConnectorClosestTo(m_primerPipe_1, m_sub1Ints_pnt);

                        tempVector = newLocationCurve.Direction;
                    }
                    else
                    {
                        // Connect to sub pipe
                        var line = Line.CreateBound(m_sub1_pnt0, m_sub1_pnt1);
                        (FirstPipe.Location as LocationCurve).Curve = line;
                        c3 = Common.GetConnectorClosestTo(FirstPipe, m_sub1Ints_pnt);

                        tempVector = line.Direction;
                    }

                    Connector c4 = null;
                    // Create primer pipe 2
                    m_primerPipe_2 = null;

                    if (g(Diameter, SecondPipe.Diameter) == false)
                    {
                        // Create primer pipe 1
                        m_primerPipe_2 = Common.Clone(VerticalPipe) as Pipe;
                        Line newLocationCurve = Line.CreateBound(m_sub2_pnt0, m_sub2_pnt1);

                        XYZ p1 = m_sub2_pnt0;

                        if (m_sub2_pnt0.DistanceTo(m_sub2Ints_pnt) < 0.01)
                        {
                            p1 = m_sub2_pnt1;
                        }

                        (m_primerPipe_2.Location as LocationCurve).Curve = Line.CreateBound(m_sub2Ints_pnt, p1);

                        c4 = Common.GetConnectorClosestTo(m_primerPipe_2, m_sub2Ints_pnt);
                    }
                    else
                    {
                        // Connect to sub pipe
                        var line = Line.CreateBound(m_sub2_pnt0, m_sub2_pnt1);
                        (SecondPipe.Location as LocationCurve).Curve = line;
                        c4 = Common.GetConnectorClosestTo(FirstPipe, m_sub1Ints_pnt);

                        tempVector = line.Direction;
                    }

                    // Create top tee
                    m_topTee = CreatTee(c3, c4, cntBottom_topTee);

                    if (m_topTee == null)
                    {
                        return false;
                    }

                    Connector cntTopTee_main1 = null;
                    Connector cntTopTee_main2 = null;
                    Connector cntTopTee_bottom = null;
                    Common.GetInfo(m_topTee, tempVector, out cntTopTee_main1, out cntTopTee_main2, out cntTopTee_bottom);

                    Pipe checkPrimerTemp1 = GetPipeConnect(cntTopTee_main1.AllRefs);
                    Pipe checkPrimerTemp2 = GetPipeConnect(cntTopTee_main2.AllRefs);

                    XYZ main1_p = cntTopTee_main1.Origin;
                    XYZ v_m_1 = cntTopTee_main1.CoordinateSystem.BasisZ;

                    var main2_p = cntTopTee_main2.Origin;
                    var v_m_2 = cntTopTee_main2.CoordinateSystem.BasisZ;

                    // Pipe 1
                    HandlerConnectTopReducer(FirstPipe,
                                             m_primerPipe_1,
                                             checkPrimerTemp1,
                                             checkPrimerTemp2,
                                             main1_p,
                                             v_m_1,
                                             main2_p,
                                             v_m_2,
                                             diameter_10,
                                             m_sub1Ints_pnt,
                                             m_sub1_pnt0,
                                             m_sub1_pnt1,
                                             m_topTee,
                                             out m_topReducer1);

                    // Pipe 2
                    HandlerConnectTopReducer(SecondPipe,
                                             m_primerPipe_2,
                                             checkPrimerTemp1,
                                             checkPrimerTemp2,
                                             main1_p,
                                             v_m_1,
                                             main2_p,
                                             v_m_2,
                                             diameter_10,
                                             m_sub2Ints_pnt,
                                             m_sub2_pnt0,
                                             m_sub2_pnt1,
                                             m_topTee,
                                             out m_topReducer2);

                    if (FlagAddNipple == true && NippleFamily != null)
                    {
                        // Nipple Family
                        HandlerConnectAccessoryNipple(FlagAddNipple, m_primerPipe_1, m_topReducer1, m_topTee);
                        HandlerConnectAccessoryNipple(FlagAddNipple, m_primerPipe_2, m_topReducer2, m_topTee);
                    }

                    return true;
                }
            }
            catch (Exception)
            { }
            return false;
        }

        public bool HandleAccessoryPairing()
        {
            try
            {
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
        private bool HandlerDivide(Pipe mainPipe, Pipe subPipe, bool flagTwoPipe, out XYZ intsMain_pnt, out XYZ sub_pnt0, out XYZ sub_pnt1, out XYZ subInts_pnt)
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

        /// <summary>
        /// Get Preferred Junction Type
        /// </summary>
        /// <param name="pipe"></param>
        /// <returns></returns>
        private PreferredJunctionType GetPreferredJunctionType(Pipe pipe)
        {
            var pipeType = pipe.PipeType as PipeType;

            return pipeType.RoutingPreferenceManager.PreferredJunctionType;
        }

        /// <summary>
        /// Create Tee Fitting
        /// </summary>
        /// <param name="pipeMain"></param>
        /// <param name="pipeCurrent"></param>
        /// <param name="splitPoint"></param>
        /// <param name="main2"></param>
        /// <returns></returns>
        public FamilyInstance CreateTeeFitting(Pipe pipeMain, Pipe pipeCurrent, XYZ splitPoint, out Pipe main2)
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

        /// <summary>
        /// Create Elbow Fitting
        /// </summary>
        /// <param name="pipeMain"></param>
        /// <param name="pipeCurrent"></param>
        /// <param name="splitPoint"></param>
        /// <param name="main2"></param>
        /// <returns></returns>
        public FamilyInstance CreateElbowFitting(Pipe pipeMain, Pipe pipeCurrent, XYZ splitPoint, out Pipe main2)
        {
            main2 = null;
            try
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
                    FamilyInstance retFitting = null;
                    if (main2.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() < pipeTempMain1.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble())
                    {
                        retFitting = Global.UIDoc.Document.Create.NewElbowFitting(c3, c5);

                        Global.UIDoc.Document.Delete(main2.Id);
                        //main2 = null;
                    }
                    else
                    {
                        retFitting = Global.UIDoc.Document.Create.NewElbowFitting(c4, c5);
                        Global.UIDoc.Document.Delete(pipeTempMain1.Id);
                    }
                    return retFitting;
                }
                catch (Exception)
                {
                    throw;
                }
            }
            catch (Exception)
            { }
            return null;
        }

        /// <summary>
        /// Get Connector Closest To
        /// </summary>
        /// <param name="e"></param>
        /// <param name="p"></param>
        /// <returns></returns>
        public Connector GetConnectorClosestTo(Element e,
                                                      XYZ p)
        {
            ConnectorManager cm = GetConnectorManager(e);

            return null == cm
              ? null
              : GetConnectorClosestTo(cm.Connectors, p);
        }

        /// <summary>
        /// Get Connector Closest To
        /// </summary>
        /// <param name="connectors"></param>
        /// <param name="p"></param>
        /// <returns></returns>
        private Connector GetConnectorClosestTo(ConnectorSet connectors,
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

        /// <summary>
        /// GetConnectorManager
        /// </summary>
        /// <param name="e"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentException"></exception>
        private ConnectorManager GetConnectorManager(Element e)
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

        public bool se(MEPCurve mepCurveSplit1, MEPCurve mepCurveSplit2)
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

        public FamilyInstance CreatTee(Connector c3, Connector c4, Connector c5)
        {
            if (c3 == null || c4 == null || c5 == null)
                return null;
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

        private bool g(double d1, double d2)
        {
            if (Math.Abs(d1 - d2) < 0.001)
                return true;

            return false;
        }

        private Pipe GetPipeConnect(ConnectorSet cs)
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

        public bool HandlerConnectTopReducer(Pipe firstPipe,
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
            try
            {
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
            catch (Exception)
            { }
            return false;
        }

        private void SetLocationLine(Pipe pipe, XYZ pOn, XYZ pOther)
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

        public FamilyInstance CreatTransitionFitting(MEPCurve mep1, MEPCurve mep2, bool checkDistance = false)
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

        public bool iss(MEPCurve mep1, MEPCurve mep2)
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

        public bool HandlerConnectAccessoryNipple(bool addNipple, Pipe primerPipe_1, FamilyInstance reducer_1, FamilyInstance tee)
        {
            try
            {
                if (addNipple)
                {
                    Tuple<Connector, Connector> cntsNipple_1 = PipeToPairConnector(primerPipe_1, reducer_1, tee);

                    if (cntsNipple_1.Item1 == null || cntsNipple_1.Item2 == null)
                        return false;

                    Line lineNipple = Line.CreateBound(cntsNipple_1.Item1.Origin, cntsNipple_1.Item2.Origin);

                    XYZ center = MiddlePoint(cntsNipple_1.Item1.Origin, cntsNipple_1.Item2.Origin);

                    FamilyInstance nipple = Global.UIDoc.Document.Create.NewFamilyInstance(center, NippleFamily, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

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
                    temp3_tee.Remove(temp3_tee.OrderBy(item => item.Origin.Z).FirstOrDefault());
                    List<Connector> temp3_reducer = GetConnectors(reducer_1.MEPModel.ConnectorManager.Connectors, true);

                    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                    // Error
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

                    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                }
            }
            catch (Exception)
            { }
            return false;
        }

        public Tuple<Connector, Connector> PipeToPairConnector(Pipe pipe, FamilyInstance reducer, FamilyInstance tee)
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

        private List<Connector> GetConnectors(ConnectorSet connectorSet, bool filter = false)
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

        public XYZ MiddlePoint(XYZ xYZ_1, XYZ xYZ_2) => new XYZ((xYZ_1.X + xYZ_2.X) * 0.5, (xYZ_1.Y + xYZ_2.Y) * 0.5, (xYZ_1.Z + xYZ_2.Z) * 0.5);

        public void RotateLine(Document doc, FamilyInstance wye, Line axisLine)
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

        public XYZ GetUnBoundIntersection(Line Line1, Line Line2)
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

        public bool SetBuiltinParameterValue(Element elem, BuiltInParameter paramId, object value)
        {
            Parameter param = elem.get_Parameter(paramId);
            return SetParameterValue(param, value);
        }

        /// <summary>
        /// Set value to parameter based on its storage type
        /// </summary>
        private bool SetParameterValue(Parameter param, object value)
        {
            if (param != null
                && !param.IsReadOnly
                && value != null)
            {
                try
                {
                    switch (param.StorageType)
                    {
                        case StorageType.Integer:
                            param.Set((int)value);
                            break;

                        case StorageType.Double:
                            param.Set((double)value);
                            break;

                        case StorageType.String:
                            param.Set((string)value);
                            break;

                        case StorageType.ElementId:
                            param.Set((ElementId)value);
                            break;
                    }
                    return true;
                }
                catch (Exception) { }
            }
            return false;
        }

        /// <summary>
        /// Get builtin parameter value based on its storage type
        /// </summary>
        public dynamic GetBuiltInParameterValue(Element elem, BuiltInParameter paramId)
        {
            if (elem != null)
            {
                Parameter parameter = elem.get_Parameter(paramId);
                return GetParameterValue(parameter);
            }
            return null;
        }

        /// <summary>
        /// get parameter value based on its storage type
        /// </summary>
        private dynamic GetParameterValue(Parameter parameter)
        {
            if (parameter != null && parameter.HasValue)
            {
                switch (parameter.StorageType)
                {
                    case StorageType.Double:
                        return parameter.AsDouble();

                    case StorageType.ElementId:
                        return parameter.AsElementId();

                    case StorageType.Integer:
                        return parameter.AsInteger();

                    case StorageType.String:
                        return parameter.AsString();
                }
            }
            return null;
        }

        /// <summary>
        /// Get parameter value by name
        /// </summary>
        public dynamic GetParameterValueByName(Element elem, string paramName)
        {
            if (elem != null)
            {
                Parameter parameter = elem.LookupParameter(paramName);
                return GetParameterValue(parameter);
            }
            return null;
        }

        #endregion Method
    }
}