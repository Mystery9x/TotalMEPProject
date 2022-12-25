using Autodesk.Revit.Attributes;

using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Forms;

using TotalMEPProject.UI.FireFightingUI;
using TotalMEPProject.Ultis;
using static System.Windows.Forms.DataFormats;

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
            try
            {
                if (App._2LevelSmartForm != null && Common.IsFormSameOpen(App._2LevelSmartForm.Name))
                    App._2LevelSmartForm.Hide();
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

                TransactionGroup t = new TransactionGroup(Global.UIDoc.Document, "Two Level Smart");
                t.Start();

                var pipeType = Global.UIDoc.Document.GetElement(App._2LevelSmartForm.FamilyType) as PipeType;

                if (App._2LevelSmartForm.OptionAddElbowConnection)
                {
                    List<Pipe> mainPipes = m_mainPipes.Select(item => Global.UIDoc.Document.GetElement(item) as Pipe).ToList();
                    foreach (Pipe pipe in mainPipes)
                    {
                        PlaceBottomElbowData placeBottomElbowData = new PlaceBottomElbowData(pipe, pairs, pipeType);
                    }
                }
                else
                {
                    foreach (PairPipes pair in pairs)
                    {
                        var pipe1 = pair._Pipe1;
                        var pipe2 = pair._Pipe2;
                        SourcePipesData sourcePipesData = new SourcePipesData(m_mainPipes,
                                                                              pipe1,
                                                                              pipe2,
                                                                              pipeType,
                                                                              diameter,
                                                                              App._2LevelSmartForm.OptionAddNipple,
                                                                              App._2LevelSmartForm.SelectedNippleFamily);
                    }
                }

                t.Assimilate();

                return Result.Succeeded;
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                if (App._2LevelSmartForm != null && Common.IsFormSameOpen(App._2LevelSmartForm.Name))
                {
                    App._2LevelSmartForm.TopMost = true;
                    App._2LevelSmartForm.Show();
                    Global.UIDoc.RefreshActiveView();
                }
            }
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

        public static bool IsFormSameOpen(string nameForm)
        {
            bool checkIsOpen = false;

            foreach (System.Windows.Forms.Form openedForm in Application.OpenForms)
            {
                if (openedForm.GetType().Name == nameForm)
                {
                    openedForm.WindowState = FormWindowState.Normal;
                    checkIsOpen = true;
                    break;
                }
            }

            return checkIsOpen;
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

        private bool m_flagTopElbowConnect = false;
        private FamilyInstance m_topElbow = null;

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

        public Pipe VerticalPipe { get => m_verticalPipe; set => m_verticalPipe = value; }

        #endregion Properties

        #region Constructor

        public SourcePipesData(List<ElementId> mainPipes, Pipe firstPipe, Pipe secondPipe, PipeType pipeType, double pipeDiameter, bool flagAddNipple, FamilySymbol nippleFamily)
        {
            if (mainPipes != null)
                MainPipes = mainPipes;

            FirstPipe = firstPipe;
            SecondPipe = secondPipe;
            PipeTypeProcess = pipeType;
            Diameter = pipeDiameter;
            FlagAddNipple = flagAddNipple;
            NippleFamily = nippleFamily;

            Initialize();
        }

        #endregion Constructor

        #region Method

        private void Initialize()
        {
            using (Transaction reTrans = new Transaction(Global.UIDoc.Document, "STEP"))
            {
                try
                {
                    reTrans.Start();

                    if (BeforeProcess() == false)
                    {
                        reTrans.RollBack();
                        return;
                    }
                    if (HandleVerticalPipe() == false)
                    {
                        reTrans.RollBack();
                        return;
                    }

                    if (HandlerBottomTee() == false)
                    {
                        reTrans.RollBack();
                        return;
                    }

                    if (HandlerTopTee() == false)
                    {
                        reTrans.RollBack();
                        return;
                    }

                    reTrans.Commit();
                }
                catch (Exception)
                {
                    reTrans.RollBack();
                }
            }
        }

        private bool BeforeProcess()
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

                    SourceRotatePipe sourceRotatePipe = new SourceRotatePipe(pipeLoop, FirstPipe, SecondPipe);
                    if (sourceRotatePipe.IsValid)
                    {
                        if (sourceRotatePipe.FlagExtendPipe)
                        {
                            (pipeLoop.Location as LocationCurve).Curve = sourceRotatePipe.NewCurveMainPipe;
                        }
                        if (sourceRotatePipe.FlagCreateElbow)
                        {
                            (FirstPipe.Location as LocationCurve).Curve = sourceRotatePipe.NewCurveFirstPipe;
                            m_flagTopElbowConnect = true;
                            MainPipe = pipeLoop;
                            break;
                        }
                        else
                        {
                            if (HandlerDivide(pipeLoop, FirstPipe, flagTwoPipe, out m_intsMain1_pnt, out m_sub1_pnt0, out m_sub1_pnt1, out m_sub1Ints_pnt) == true)
                            {
                                MainPipe = pipeLoop;
                                break;
                            }
                        }
                    }
                }

                // Rotate sub pipe by main pipe
                //HandleBranchPipePerpendiculerMainPipe();

                if (m_flagTopElbowConnect == false)
                {
                    if (HandlerDivide(MainPipe, FirstPipe, flagTwoPipe, out m_intsMain1_pnt, out m_sub1_pnt0, out m_sub1_pnt1, out m_sub1Ints_pnt) == true)
                    {
                        // If only 1 pipe
                        if (SecondPipe == null)
                        {
                            (FirstPipe.Location as LocationCurve).Curve = Line.CreateBound(m_sub1_pnt0, m_sub1Ints_pnt);

                            SecondPipe = Common.Clone(FirstPipe) as Pipe;
                            (SecondPipe.Location as LocationCurve).Curve = Line.CreateBound(m_sub1Ints_pnt, m_sub1_pnt1);

                            // Set point value again
                            m_sub1_pnt1 = m_sub1Ints_pnt;
                        }

                        if (SecondPipe != null)
                        {
                            if (HandlerDivide(MainPipe, SecondPipe, true, out m_intsMain2_pnt, out m_sub2_pnt0, out m_sub2_pnt1, out m_sub2Ints_pnt) == false)
                            {
                                return false;
                            }

                            if (m_intsMain1_pnt.DistanceTo(m_intsMain2_pnt) > 0.001)
                            {
                                return false;
                            }

                            if (m_sub1Ints_pnt.DistanceTo(m_sub2Ints_pnt) > 0.001)
                            {
                                return false;
                            }
                        }
                    }
                }
                else
                {
                    return HandlerDivideOnlyPipe(MainPipe, FirstPipe, out m_intsMain1_pnt, out m_sub1_pnt0, out m_sub1_pnt1, out m_sub1Ints_pnt);
                }

                return true;
            }
            catch (Exception)
            { }
            return false;
        }

        private bool HandleVerticalPipe()
        {
            try
            {
                if (m_flagTopElbowConnect == true)
                {
                    Diameter = (double)GetParameterValueByName(FirstPipe, "Diameter");
                }

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

        private bool HandleBranchPipePerpendiculerMainPipe()
        {
            try
            {
                // MainPipe
                if (MainPipe == null)
                    return false;
                var mainPipe_curve_reality = MainPipe.GetCurve();

                XYZ mainPipe_start_pnt = mainPipe_curve_reality.GetEndPoint(0);
                XYZ mainPipe_end_pnt = mainPipe_curve_reality.GetEndPoint(1);

                XYZ mainPipe_start_pnt_2d = Common.ToPoint2D(mainPipe_curve_reality.GetEndPoint(0));
                XYZ mainPipe_end_pnt_2d = Common.ToPoint2D(mainPipe_curve_reality.GetEndPoint(1));
                XYZ mainPipe_dir_2d = (mainPipe_end_pnt_2d - mainPipe_start_pnt_2d);

                var mainPipe_curve_2d_Unbound = Line.CreateUnbound(mainPipe_start_pnt_2d, mainPipe_dir_2d);

                XYZ mainPipe_dir_2d_rotate = mainPipe_curve_2d_Unbound.Direction.Normalize().CrossProduct(XYZ.BasisZ).Normalize();

                // First Pipe
                if (FirstPipe == null)
                    return false;
                var firstpipe_curve_reality = FirstPipe.GetCurve();

                XYZ firstpipe_start_pnt = firstpipe_curve_reality.GetEndPoint(0);
                XYZ firstpipe_end_pnt = firstpipe_curve_reality.GetEndPoint(1);

                XYZ firstpipe_start_pnt_2d = Common.ToPoint2D(firstpipe_curve_reality.GetEndPoint(0));
                XYZ firstpipe_end_pnt_2d = Common.ToPoint2D(firstpipe_curve_reality.GetEndPoint(1));
                XYZ firstpipe_dir_2d = (firstpipe_end_pnt_2d - firstpipe_start_pnt_2d);

                var firstpipe_curve_2d_Unbound = Line.CreateUnbound(firstpipe_start_pnt_2d, firstpipe_dir_2d);

                List<Element> allComponentConnectFirstPipe = AllComponentsOnPipeTruss(Global.UIDoc.Document, FirstPipe.Id, true);

                XYZ ints_pnt_firstPipe_mainPipe_2d = null;

                int isLeft_1 = -1;
                int isLeft_2 = -1;

                // Second Pipe

                Curve secondpipe_curve_reality = null;

                XYZ secondpipe_start_pnt = null;
                XYZ secondpipe_end_pnt = null;

                XYZ secondpipe_start_pnt_2d = null;
                XYZ secondpipe_end_pnt_2d = null;
                XYZ secondpipe_dir_2d = null;

                Curve secondpipe_curve_2d_Unbound = null;

                // Check intersect main pipe with first pipe
                IntersectionResultArray intRetArr = new IntersectionResultArray();
                SetComparisonResult setCompResult = mainPipe_curve_2d_Unbound.Intersect(firstpipe_curve_2d_Unbound, out intRetArr);
                if (setCompResult != SetComparisonResult.Overlap)
                    return false;

                ints_pnt_firstPipe_mainPipe_2d = intRetArr.get_Item(0).XYZPoint;
                var after_firstpipe_curve_rotate = Line.CreateBound(ints_pnt_firstPipe_mainPipe_2d, ints_pnt_firstPipe_mainPipe_2d + mainPipe_dir_2d_rotate * 100);

                var after_firstpipe_curve_case1 = Line.CreateBound(ints_pnt_firstPipe_mainPipe_2d, firstpipe_start_pnt_2d + (firstpipe_start_pnt_2d - firstpipe_end_pnt_2d).Normalize() * 10);
                var after_firstpipe_curve_case2 = Line.CreateBound(ints_pnt_firstPipe_mainPipe_2d, firstpipe_end_pnt_2d - (firstpipe_start_pnt_2d - firstpipe_end_pnt_2d).Normalize() * 10);

                double rotate_angle_firstPipe_1 = after_firstpipe_curve_rotate.Direction.AngleTo((after_firstpipe_curve_case1 as Line).Direction);
                double rotate_angle_firstPipe_2 = after_firstpipe_curve_rotate.Direction.AngleTo((after_firstpipe_curve_case2 as Line).Direction);

                isLeft_1 = IsLeftMath(mainPipe_start_pnt_2d, mainPipe_end_pnt_2d, firstpipe_start_pnt_2d + (firstpipe_start_pnt_2d - firstpipe_end_pnt_2d).Normalize() * 10);
                isLeft_2 = IsLeftMath(mainPipe_start_pnt_2d, mainPipe_end_pnt_2d, firstpipe_end_pnt_2d - (firstpipe_start_pnt_2d - firstpipe_end_pnt_2d).Normalize() * 10);

                try
                {
                    bool flag = false;
                    using (SubTransaction reSubTrans = new SubTransaction(Global.UIDoc.Document))
                    {
                        reSubTrans.Start();
                        ElementTransformUtils.RotateElements(Global.UIDoc.Document, allComponentConnectFirstPipe.Select(item => item.Id).ToList(), Line.CreateUnbound(ints_pnt_firstPipe_mainPipe_2d, XYZ.BasisZ), rotate_angle_firstPipe_1);
                        Global.UIDoc.Document.Regenerate();
                        var curve_sub_Temp = FirstPipe.GetCurve();
                        var curve_sub_2d_Temp = Line.CreateUnbound(Common.ToPoint2D(curve_sub_Temp.GetEndPoint(0)), Common.ToPoint2D(curve_sub_Temp.GetEndPoint(1)) - Common.ToPoint2D(curve_sub_Temp.GetEndPoint(0)));
                        double angle = (curve_sub_2d_Temp as Line).Direction.AngleTo(mainPipe_dir_2d);
                        if (Math.Abs((angle / (Math.PI) - 0.5)) <= 0.000000001)
                        {
                            firstpipe_curve_reality = FirstPipe.GetCurve();

                            firstpipe_start_pnt = firstpipe_curve_reality.GetEndPoint(0);
                            firstpipe_end_pnt = firstpipe_curve_reality.GetEndPoint(1);
                            firstpipe_start_pnt_2d = Common.ToPoint2D(firstpipe_curve_reality.GetEndPoint(0));
                            firstpipe_end_pnt_2d = Common.ToPoint2D(firstpipe_curve_reality.GetEndPoint(1));

                            if (isLeft_1 == IsLeftMath(mainPipe_start_pnt_2d, mainPipe_end_pnt_2d, firstpipe_start_pnt_2d + (firstpipe_start_pnt_2d - firstpipe_end_pnt_2d).Normalize() * 10))
                            {
                                if (SecondPipe != null)
                                {
                                    List<Element> allComponentConnectSecondPipe = AllComponentsOnPipeTruss(Global.UIDoc.Document, SecondPipe.Id, true);
                                    ElementTransformUtils.RotateElements(Global.UIDoc.Document, allComponentConnectSecondPipe.Select(item => item.Id).ToList(), Line.CreateUnbound(ints_pnt_firstPipe_mainPipe_2d, XYZ.BasisZ), rotate_angle_firstPipe_1);
                                    Global.UIDoc.Document.Regenerate();
                                }
                            }
                            else
                            {
                                ElementTransformUtils.RotateElements(Global.UIDoc.Document, allComponentConnectFirstPipe.Select(item => item.Id).ToList(), Line.CreateUnbound(ints_pnt_firstPipe_mainPipe_2d, XYZ.BasisZ), -Math.PI);
                                Global.UIDoc.Document.Regenerate();
                                if (SecondPipe != null)
                                {
                                    List<Element> allComponentConnectSecondPipe = AllComponentsOnPipeTruss(Global.UIDoc.Document, SecondPipe.Id, true);
                                    ElementTransformUtils.RotateElements(Global.UIDoc.Document, allComponentConnectSecondPipe.Select(item => item.Id).ToList(), Line.CreateUnbound(ints_pnt_firstPipe_mainPipe_2d, XYZ.BasisZ), rotate_angle_firstPipe_1 - Math.PI);
                                    Global.UIDoc.Document.Regenerate();
                                }
                            }

                            reSubTrans.Commit();
                        }
                        else
                        {
                            flag = true;
                            reSubTrans.RollBack();
                        }
                    }

                    if (flag == true)
                    {
                        using (SubTransaction reSubTrans = new SubTransaction(Global.UIDoc.Document))
                        {
                            reSubTrans.Start();
                            ElementTransformUtils.RotateElements(Global.UIDoc.Document, allComponentConnectFirstPipe.Select(item => item.Id).ToList(), Line.CreateUnbound(ints_pnt_firstPipe_mainPipe_2d, XYZ.BasisZ), rotate_angle_firstPipe_2);
                            var curve_sub_Temp = FirstPipe.GetCurve();
                            var curve_sub_2d_Temp = Line.CreateUnbound(Common.ToPoint2D(curve_sub_Temp.GetEndPoint(0)), Common.ToPoint2D(curve_sub_Temp.GetEndPoint(1)) - Common.ToPoint2D(curve_sub_Temp.GetEndPoint(0)));
                            double angle = (curve_sub_2d_Temp as Line).Direction.AngleTo(mainPipe_dir_2d);
                            firstpipe_curve_reality = FirstPipe.GetCurve();

                            firstpipe_start_pnt = firstpipe_curve_reality.GetEndPoint(0);
                            firstpipe_end_pnt = firstpipe_curve_reality.GetEndPoint(1);
                            firstpipe_start_pnt_2d = Common.ToPoint2D(firstpipe_curve_reality.GetEndPoint(0));
                            firstpipe_end_pnt_2d = Common.ToPoint2D(firstpipe_curve_reality.GetEndPoint(1));

                            if (isLeft_1 == IsLeftMath(mainPipe_start_pnt_2d, mainPipe_end_pnt_2d, firstpipe_start_pnt_2d + (firstpipe_start_pnt_2d - firstpipe_end_pnt_2d).Normalize() * 10))
                            {
                                if (SecondPipe != null)
                                {
                                    List<Element> allComponentConnectSecondPipe = AllComponentsOnPipeTruss(Global.UIDoc.Document, SecondPipe.Id, true);
                                    ElementTransformUtils.RotateElements(Global.UIDoc.Document, allComponentConnectSecondPipe.Select(item => item.Id).ToList(), Line.CreateUnbound(ints_pnt_firstPipe_mainPipe_2d, XYZ.BasisZ), rotate_angle_firstPipe_2);
                                }
                            }
                            else
                            {
                                ElementTransformUtils.RotateElements(Global.UIDoc.Document, allComponentConnectFirstPipe.Select(item => item.Id).ToList(), Line.CreateUnbound(ints_pnt_firstPipe_mainPipe_2d, XYZ.BasisZ), -Math.PI);
                                if (SecondPipe != null)
                                {
                                    List<Element> allComponentConnectSecondPipe = AllComponentsOnPipeTruss(Global.UIDoc.Document, SecondPipe.Id, true);
                                    ElementTransformUtils.RotateElements(Global.UIDoc.Document, allComponentConnectSecondPipe.Select(item => item.Id).ToList(), Line.CreateUnbound(ints_pnt_firstPipe_mainPipe_2d, XYZ.BasisZ), rotate_angle_firstPipe_2 - Math.PI);
                                }
                            }

                            reSubTrans.Commit();
                        }
                    }
                }
                catch (Exception)
                { }

                return true;
            }
            catch (Exception)
            { }
            return false;
        }

        public int IsLeftMath(XYZ Startpoint, XYZ Endpoint, XYZ P)
        {
            double Ans = ((Endpoint.X - Startpoint.X) * (P.Y - Startpoint.Y) -
              (P.X - Startpoint.X) * (Endpoint.Y - Startpoint.Y));
            if (Math.Abs(Ans) < 1.0e-8)
            { return 0; } //P is on the line
            else
            {
                if (Ans > 0)
                { return 1; } //P is left of the line (CW)
                else
                { return -1; } //P is right of the line (CCW)
            }
        }

        private bool HandlerBottomTee(bool flag = true)
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
                            if (flag == true)
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
                    return se(MainPipe, VerticalPipe);
                }
            }
            catch (Exception)
            { }
            return false;
        }

        private bool HandlerTopTee()
        {
            try
            {
                // Process with sub pipes

                Connector cntBottom_topTee = GetConnectorClosestTo(VerticalPipe, m_sub1Ints_pnt);

                //double diameter_primer = 15 * Common.mmToFT;

                double diameter_10 = 10 * Common.mmToFT;

                if (m_flagTopElbowConnect == true)
                {
                    (FirstPipe.Location as LocationCurve).Curve = Line.CreateBound(m_sub1_pnt0, m_sub1_pnt1);
                    //Connect
                    Connector c3 = GetConnectorClosestTo(FirstPipe, m_sub1Ints_pnt);
                    m_topElbow = CreateElbow(c3, cntBottom_topTee);
                    if (m_topElbow == null)
                    {
                        return false;
                    }
                    return true;
                }
                else
                {
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

                        if (m_flagTopElbowConnect == true)
                        {
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
            Global.UIDoc.Document.Regenerate();
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
        private bool HandlerDivideOnlyPipe(Pipe mainPipe, Pipe subPipe, out XYZ intsMain_pnt, out XYZ sub_pnt0, out XYZ sub_pnt1, out XYZ subInts_pnt)
        {
            Global.UIDoc.Document.Regenerate();
            intsMain_pnt = null;
            sub_pnt0 = null;
            sub_pnt1 = null;
            subInts_pnt = null;

            var curve_main = mainPipe.GetCurve();
            var curve_main_2d = Line.CreateBound(Common.ToPoint2D(curve_main.GetEndPoint(0)), Common.ToPoint2D(curve_main.GetEndPoint(1)));

            var curve_sub = subPipe.GetCurve();

            Curve curve2d = null;
            Curve curve3d = null;

            curve2d = Line.CreateBound(Common.ToPoint2D(curve_sub.GetEndPoint(0)), Common.ToPoint2D(curve_sub.GetEndPoint(1)));
            curve3d = curve_sub;

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

            sub_pnt0 = curve_sub.GetEndPoint(0);
            sub_pnt1 = curve_sub.GetEndPoint(1);

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
        private FamilyInstance CreateTeeFitting(Pipe pipeMain, Pipe pipeCurrent, XYZ splitPoint, out Pipe main2)
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
        private FamilyInstance CreateElbowFitting(Pipe pipeMain, Pipe pipeCurrent, XYZ splitPoint, out Pipe main2)
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
        private Connector GetConnectorClosestTo(Element e,
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

        /// <summary>
        /// s e
        /// </summary>
        /// <param name="mepCurveSplit1"></param>
        /// <param name="mepCurveSplit2"></param>
        /// <returns></returns>
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

        /// <summary>
        /// Create Tee
        /// </summary>
        /// <param name="c3"></param>
        /// <param name="c4"></param>
        /// <param name="c5"></param>
        /// <returns></returns>
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

        /// <summary>
        /// Create Tee
        /// </summary>
        /// <param name="c3"></param>
        /// <param name="c4"></param>
        /// <param name="c5"></param>
        /// <returns></returns>
        public FamilyInstance CreateElbow(Connector c3, Connector c4)
        {
            if (c3 == null || c4 == null)
                return null;
            try
            {
                var fitting = Global.UIDoc.Document.Create.NewElbowFitting(c3, c4);

                return fitting;
            }
            catch (System.Exception ex)
            {
                return null;
            }
        }

        /// <summary>
        /// g
        /// </summary>
        /// <param name="d1"></param>
        /// <param name="d2"></param>
        /// <returns></returns>
        private bool g(double d1, double d2)
        {
            if (Math.Abs(d1 - d2) < 0.001)
                return true;

            return false;
        }

        /// <summary>
        /// Get pipe connect
        /// </summary>
        /// <param name="cs"></param>
        /// <returns></returns>
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

        /// <summary>
        /// Handler Connect Top Reducer
        /// </summary>
        /// <param name="firstPipe"></param>
        /// <param name="primerPipe_1"></param>
        /// <param name="checkPrimerTemp1"></param>
        /// <param name="checkPrimerTemp2"></param>
        /// <param name="main1_p"></param>
        /// <param name="v_m_1"></param>
        /// <param name="main2_p"></param>
        /// <param name="v_m_2"></param>
        /// <param name="diameter_10"></param>
        /// <param name="sub1Ints_pnt"></param>
        /// <param name="sub1_pnt0"></param>
        /// <param name="sub1_pnt1"></param>
        /// <param name="tee"></param>
        /// <param name="reducer_1"></param>
        /// <returns></returns>
        private bool HandlerConnectTopReducer(Pipe firstPipe,
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

        /// <summary>
        /// Set Location Line
        /// </summary>
        /// <param name="pipe"></param>
        /// <param name="pOn"></param>
        /// <param name="pOther"></param>
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

        /// <summary>
        /// Creat Transition Fitting
        /// </summary>
        /// <param name="mep1"></param>
        /// <param name="mep2"></param>
        /// <param name="checkDistance"></param>
        /// <returns></returns>
        private FamilyInstance CreatTransitionFitting(MEPCurve mep1, MEPCurve mep2, bool checkDistance = false)
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

        /// <summary>
        /// Is Same
        /// </summary>
        /// <param name="mep1"></param>
        /// <param name="mep2"></param>
        /// <returns></returns>
        private bool iss(MEPCurve mep1, MEPCurve mep2)
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

        /// <summary>
        /// Handler Connect Accessory Nipple
        /// </summary>
        /// <param name="addNipple"></param>
        /// <param name="primerPipe_1"></param>
        /// <param name="reducer_1"></param>
        /// <param name="tee"></param>
        /// <returns></returns>
        private bool HandlerConnectAccessoryNipple(bool addNipple, Pipe primerPipe_1, FamilyInstance reducer_1, FamilyInstance tee)
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

                    List<Connector> temp3 = GetConnectors(nipple.MEPModel.ConnectorManager.Connectors, true).Where(item => item.IsConnected == false).ToList();
                    List<Connector> temp3_tee = GetConnectors(tee.MEPModel.ConnectorManager.Connectors, true).ToList();
                    temp3_tee.Remove(temp3_tee.OrderBy(item => item.Origin.Z).FirstOrDefault());
                    temp3_tee = temp3_tee.Where(item => item.IsConnected == false).ToList();
                    List<Connector> temp3_reducer = GetConnectors(reducer_1.MEPModel.ConnectorManager.Connectors, true).Where(item => item.IsConnected == false).ToList();

                    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                    //// Error

                    foreach (Connector cnt in temp3_reducer)
                    {
                        bool flag = false;
                        foreach (Connector cnt1 in temp3)
                        {
                            if (Common.IsEqual(cnt1.Origin, cnt.Origin) && cnt.IsConnected == false)
                            {
                                cnt1.ConnectTo(cnt);
                                flag = true;
                            }
                        }
                        if (flag == true)
                            break;
                    }

                    foreach (Connector cnt in temp3_tee)
                    {
                        bool flag = false;
                        foreach (Connector cnt1 in temp3)
                        {
                            if (Common.IsEqual(cnt1.Origin, cnt.Origin) && cnt.IsConnected == false)
                            {
                                cnt1.ConnectTo(cnt);
                                flag = true;
                            }
                        }
                        if (flag == true)
                            break;
                    }

                    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                }
            }
            catch (Exception)
            { }
            return false;
        }

        /// <summary>
        /// Pipe To Pair Connector
        /// </summary>
        /// <param name="pipe"></param>
        /// <param name="reducer"></param>
        /// <param name="tee"></param>
        /// <returns></returns>
        private Tuple<Connector, Connector> PipeToPairConnector(Pipe pipe, FamilyInstance reducer, FamilyInstance tee)
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

        private XYZ MiddlePoint(XYZ xYZ_1, XYZ xYZ_2) => new XYZ((xYZ_1.X + xYZ_2.X) * 0.5, (xYZ_1.Y + xYZ_2.Y) * 0.5, (xYZ_1.Z + xYZ_2.Z) * 0.5);

        private void RotateLine(Document doc, FamilyInstance wye, Line axisLine)
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

        private XYZ GetUnBoundIntersection(Line Line1, Line Line2)
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

        private bool SetBuiltinParameterValue(Element elem, BuiltInParameter paramId, object value)
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
        private dynamic GetBuiltInParameterValue(Element elem, BuiltInParameter paramId)
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
        private dynamic GetParameterValueByName(Element elem, string paramName)
        {
            if (elem != null)
            {
                Parameter parameter = elem.LookupParameter(paramName);
                return GetParameterValue(parameter);
            }
            return null;
        }

        /// <summary>
        /// All Components On Pipe Truss
        /// </summary>
        /// <param name="document"></param>
        /// <param name="elementMainId"></param>
        /// <param name="addFirst"></param>
        /// <param name="rmvElement"></param>
        /// <returns></returns>
        private List<Element> AllComponentsOnPipeTruss(Autodesk.Revit.DB.Document document, ElementId elementMainId, bool addFirst = false, List<Element> rmvElement = null)
        {
            List<Element> components = new List<Element>();
            Element eleMain = document.GetElement(elementMainId);
            List<Element> rmvElement_1 = null;
            if (addFirst == true)
            {
                rmvElement_1 = new List<Element> { eleMain };
                components.Add(eleMain);
            }
            else
            {
                rmvElement_1 = new List<Element>(rmvElement);
            }
            if (eleMain != null)
            {
                List<Connector> connectors = new List<Connector>();
                if (eleMain is MEPCurve processMEPCurve && processMEPCurve.ConnectorManager != null)
                {
                    foreach (Connector connector in processMEPCurve.ConnectorManager.Connectors)
                    {
                        if (connector.ConnectorType != ConnectorType.End)
                            continue;
                        connectors.Add(connector);
                    }
                }
                else if (eleMain is FamilyInstance fmlIns && fmlIns.MEPModel.ConnectorManager != null)
                {
                    foreach (Connector connector in fmlIns.MEPModel.ConnectorManager.Connectors)
                    {
                        if (connector.ConnectorType != ConnectorType.End && connector.IsConnected == true)
                            continue;
                        connectors.Add(connector);
                    }
                }

                if (connectors.Count > 0)
                {
                    foreach (Connector cnt1 in connectors)
                    {
                        foreach (Connector cnt2 in cnt1.AllRefs)
                        {
                            Element eleCheck = cnt2.Owner;
                            if (null != eleCheck && (eleCheck is MEPCurve || eleCheck is FamilyInstance))
                            {
                                if (rmvElement_1 != null && rmvElement_1.Any(item => item.Id == eleCheck.Id))
                                    continue;
                                else
                                {
                                    components.Add(eleCheck);
                                    List<Element> rmvElement_2 = new List<Element>(rmvElement_1);
                                    rmvElement_2.Add(eleCheck);

                                    components.AddRange(AllComponentsOnPipeTruss(document, eleCheck.Id, false, rmvElement_2));
                                }
                            }
                        }
                    }
                }
            }

            return components;
        }

        #endregion Method
    }

    public class SourceRotatePipe
    {
        public Pipe MainPipe { get; set; }
        public Line CurveMainPipe { get; set; }
        public Line CurveMainPipe_2d { get; set; }
        public Pipe FirstPipe { get; set; }
        public Line CurveFirstPipe { get; set; }
        public Line CurveFirstPipe_2d { get; set; }

        public Line NewCurveFirstPipe { get; set; }

        public Line NewCurveFirstPipe_2d { get; set; }

        public Line NewCurveMainPipe { get; set; }
        public Line NewCurveMainPipe_2d { get; set; }

        public Pipe SecondPipe { get; set; }
        public Line CurveSecondPipe { get; set; }
        public Line CurveSecondPipe_2d { get; set; }
        public bool IsValid { get; set; }
        public bool FlagCreateTopTee { get; set; }
        public bool FlagCreateElbow { get; set; }

        public XYZ IntsMain { get; set; }
        public XYZ SubMain { get; set; }

        public bool FlagExtendPipe { get; set; }

        public SourceRotatePipe(Pipe mainPipe, Pipe firstPipe, Pipe secondPipe)
        {
            IsValid = false;
            if (mainPipe != null && firstPipe != null)
            {
                MainPipe = mainPipe;
                FirstPipe = firstPipe;
                SecondPipe = secondPipe;
                if (Initialize())
                    IsValid = true;
            }
        }

        public bool Initialize()
        {
            try
            {
                CurveMainPipe = MainPipe.GetCurve() as Line;
                CurveMainPipe_2d = Line.CreateBound(ToPoint2D(CurveMainPipe.GetEndPoint(0)), ToPoint2D(CurveMainPipe.GetEndPoint(1)));

                CurveFirstPipe = FirstPipe.GetCurve() as Line;
                CurveFirstPipe_2d = Line.CreateBound(ToPoint2D(CurveFirstPipe.GetEndPoint(0)), ToPoint2D(CurveFirstPipe.GetEndPoint(1)));

                if (SecondPipe != null && SecondPipe.IsValidObject)
                {
                    CurveSecondPipe = SecondPipe.GetCurve() as Line;
                    CurveSecondPipe_2d = Line.CreateBound(ToPoint2D(CurveSecondPipe.GetEndPoint(0)), ToPoint2D(CurveSecondPipe.GetEndPoint(1)));
                }

                if (FirstPipe != null && SecondPipe != null)
                {
                    if (RealityIntersect(CurveMainPipe_2d, CurveFirstPipe_2d) || RealityIntersect(CurveMainPipe_2d, CurveSecondPipe_2d))
                        return true;
                    else if (!RealityIntersect(CurveMainPipe_2d, CurveFirstPipe_2d) && !RealityIntersect(CurveMainPipe_2d, CurveSecondPipe_2d))
                    {
                        XYZ expanDirectionn = (ToPoint2D(CurveFirstPipe_2d.GetEndPoint(1)) - ToPoint2D(CurveFirstPipe_2d.GetEndPoint(0))).Normalize();

                        var expandFirstPipe = Line.CreateBound(ToPoint2D(CurveFirstPipe_2d.GetEndPoint(0)) - expanDirectionn * 400 * mmToFT, ToPoint2D(CurveFirstPipe_2d.GetEndPoint(1) + expanDirectionn * 400 * mmToFT));
                        if (RealityIntersect(expandFirstPipe, CurveMainPipe_2d))
                            return true;

                        List<Connector> cntMainPipe = new List<Connector>();
                        foreach (Connector connector in MainPipe.ConnectorManager.Connectors)
                        {
                            if (connector.ConnectorType != ConnectorType.End)
                                continue;
                            cntMainPipe.Add(connector);
                        }

                        if (cntMainPipe.Any(item => item.IsConnected == true) == false)
                        {
                            return true;
                        }
                        if (cntMainPipe.Count >= 2)
                        {
                            var directionExpandPipe_2d = ToPoint2D(cntMainPipe[1].Origin) - ToPoint2D(cntMainPipe[0].Origin);

                            if (cntMainPipe[0].IsConnected == false)
                            {
                                NewCurveMainPipe_2d = Line.CreateBound(ToPoint2D(cntMainPipe[0].Origin - directionExpandPipe_2d * 100 * mmToFT), ToPoint2D(cntMainPipe[1].Origin));
                            }
                            else if (cntMainPipe[1].IsConnected == false)
                            {
                                NewCurveMainPipe_2d = Line.CreateBound(ToPoint2D(cntMainPipe[0].Origin), ToPoint2D(cntMainPipe[1].Origin + directionExpandPipe_2d * 100 * mmToFT));
                            }
                            else
                                return true;

                            if (RealityIntersect(NewCurveMainPipe_2d, expandFirstPipe))
                            {
                                FlagExtendPipe = true;
                                var directionExpandPipe = cntMainPipe[1].Origin - cntMainPipe[0].Origin;
                                if (cntMainPipe[0].IsConnected == false)
                                {
                                    NewCurveMainPipe = Line.CreateBound(cntMainPipe[0].Origin - directionExpandPipe * 100 * mmToFT, cntMainPipe[1].Origin);
                                }
                                else if (cntMainPipe[1].IsConnected == false)
                                {
                                    NewCurveMainPipe = Line.CreateBound(cntMainPipe[0].Origin, cntMainPipe[1].Origin + directionExpandPipe * 100 * mmToFT);
                                }
                                return true;
                            }
                        }

                        return true;
                    }
                }
                else if (FirstPipe != null && SecondPipe == null)
                {
                    if (RealityIntersect(CurveMainPipe_2d, CurveFirstPipe_2d))
                        return true;
                    else
                    {
                        return SpecialProcesFirstPipe();
                    }
                }
            }
            catch (Exception)
            { }
            return false;
        }

        private bool RealityIntersect(Line mainLine, Line checkLine)
        {
            try
            {
                IntersectionResultArray intsRetArr = new IntersectionResultArray();
                var intsRet = mainLine.Intersect(checkLine, out intsRetArr);
                if (intsRet == SetComparisonResult.Overlap)
                    return true;
            }
            catch (Exception)
            { }
            return false;
        }

        public double mmToFT = 0.0032808399;

        private bool SpecialProcesFirstPipe()
        {
            try
            {
                if (FirstPipe == null)
                    return false;

                // Check direction firte pipe valid
                // Get connector
                List<Connector> cntOfPipe = new List<Connector>();
                foreach (Connector connector in FirstPipe.ConnectorManager.Connectors)
                {
                    if (connector.ConnectorType != ConnectorType.End)
                        continue;
                    if (connector.IsConnected == true)
                        continue;
                    cntOfPipe.Add(connector);
                }

                if (cntOfPipe.Count <= 0)
                    return false;

                XYZ directionExpand = (ToPoint2D(CurveFirstPipe.GetEndPoint(1)) - ToPoint2D(CurveFirstPipe.GetEndPoint(0))).Normalize();
                XYZ validContOrg = cntOfPipe[0].Origin;

                if (cntOfPipe.Count == 2 && CurveMainPipe_2d.Distance(ToPoint2D(cntOfPipe[0].Origin)) > CurveMainPipe_2d.Distance(ToPoint2D(cntOfPipe[1].Origin)))
                {
                    validContOrg = cntOfPipe[1].Origin;
                }
                bool flag = false;

                if (IsEqual(CurveFirstPipe.GetEndPoint(0), validContOrg))
                {
                    NewCurveFirstPipe_2d = Line.CreateBound(ToPoint2D(CurveFirstPipe.GetEndPoint(0)) - directionExpand * 400 * mmToFT, ToPoint2D(CurveFirstPipe.GetEndPoint(1)));
                }
                else
                {
                    NewCurveFirstPipe_2d = Line.CreateBound(ToPoint2D(CurveFirstPipe.GetEndPoint(0)), ToPoint2D(CurveFirstPipe.GetEndPoint(1)) + directionExpand * 400 * mmToFT);

                    flag = true;
                }

                if (RealityIntersect(CurveMainPipe_2d, NewCurveFirstPipe_2d))
                {
                    IntersectionResultArray intsRetArr = new IntersectionResultArray();
                    var intsRet = CurveMainPipe_2d.Intersect(NewCurveFirstPipe_2d, out intsRetArr);
                    var intsPnt_2d = intsRetArr.get_Item(0).XYZPoint;

                    var temp_Direction = Line.CreateUnbound(CurveFirstPipe.GetEndPoint(0), CurveFirstPipe.GetEndPoint(1) - CurveFirstPipe.GetEndPoint(0));
                    IntersectionResult porjectLine = temp_Direction.Project(intsPnt_2d);
                    XYZ newPnt = porjectLine.XYZPoint;

                    if (flag)
                    {
                        //Find 3d
                        double temp = 200;

                        var lineZ = Line.CreateBound(new XYZ(intsPnt_2d.X, intsPnt_2d.Y, intsPnt_2d.Z - temp), new XYZ(intsPnt_2d.X, intsPnt_2d.Y, intsPnt_2d.Z + temp));
                        var arr = new IntersectionResultArray();
                        var result = lineZ.Intersect(temp_Direction, out arr);
                        if (result != SetComparisonResult.Overlap)
                            return false;

                        newPnt = arr.get_Item(0).XYZPoint;

                        NewCurveFirstPipe = Line.CreateBound(CurveFirstPipe.GetEndPoint(0), newPnt);

                        arr = new IntersectionResultArray();
                        result = lineZ.Intersect(CurveMainPipe, out arr);
                        if (result != SetComparisonResult.Overlap)
                            return false;

                        IntsMain = arr.get_Item(0).XYZPoint;

                        //Find 3d on two sub pipe
                        arr = new IntersectionResultArray();
                        result = lineZ.Intersect(NewCurveFirstPipe, out arr);
                        if (result != SetComparisonResult.Overlap)
                            return false;

                        SubMain = arr.get_Item(0).XYZPoint;

                        FlagCreateElbow = true;
                        return true;
                    }
                    else
                    {
                        //Find 3d
                        double temp = 200;

                        var lineZ = Line.CreateBound(new XYZ(intsPnt_2d.X, intsPnt_2d.Y, intsPnt_2d.Z - temp), new XYZ(intsPnt_2d.X, intsPnt_2d.Y, intsPnt_2d.Z + temp));
                        var arr = new IntersectionResultArray();
                        var result = lineZ.Intersect(temp_Direction, out arr);
                        if (result != SetComparisonResult.Overlap)
                            return false;

                        newPnt = arr.get_Item(0).XYZPoint;

                        NewCurveFirstPipe = Line.CreateBound(newPnt, CurveFirstPipe.GetEndPoint(1));

                        arr = new IntersectionResultArray();
                        result = lineZ.Intersect(CurveMainPipe, out arr);
                        if (result != SetComparisonResult.Overlap)
                            return false;

                        IntsMain = arr.get_Item(0).XYZPoint;

                        //Find 3d on two sub pipe
                        arr = new IntersectionResultArray();
                        result = lineZ.Intersect(NewCurveFirstPipe, out arr);
                        if (result != SetComparisonResult.Overlap)
                            return false;

                        SubMain = arr.get_Item(0).XYZPoint;

                        FlagCreateElbow = true;
                        return true;
                    }
                }
                else
                {
                    // Case 2 Expand mainpipe
                    //cntOfPipe.Clear();
                    //foreach (Connector connector in FirstPipe.ConnectorManager.Connectors)
                    //{
                    //    if (connector.ConnectorType != ConnectorType.End)
                    //        continue;

                    //    cntOfPipe.Add(connector);
                    //}

                    //directionExpand = (ToPoint2D(CurveMainPipe.GetEndPoint(1)) - ToPoint2D(CurveMainPipe.GetEndPoint(0))).Normalize();

                    //if (cntOfPipe[0].IsConnected == false)
                    //{
                    //    NewCurveMainPipe = Line.CreateBound(ToPoint2D(cntOfPipe[0].Origin - directionExpand * 200 * mmToFT), ToPoint2D(cntOfPipe[1].Origin));
                    //}
                    //else if (cntOfPipe[1].IsConnected == false)
                    //{
                    //    NewCurveMainPipe = Line.CreateBound(ToPoint2D(cntOfPipe[0].Origin), ToPoint2D(cntOfPipe[1].Origin + directionExpand * 200 * mmToFT));
                    //}
                    //else
                    //    return false;

                    //if (RealityIntersect(NewCurveMainPipe, NewCurveFirstPipe_2d))
                    //{
                    //    FlagExtendPipe = true;
                    //    return true;
                    //}

                    return false;
                }
            }
            catch (Exception)
            { }
            return false;
        }

        private XYZ ToPoint2D(XYZ point3d, double z = 0)
        {
            return new XYZ(point3d.X, point3d.Y, z);
        }

        public bool IsEqual(double first, double second, double tolerance = 10e-5)
        {
            double result = Math.Abs(first - second);
            return result < tolerance;
        }

        public bool IsEqual(XYZ first, XYZ second)
        {
            return IsEqual(first.X, second.X)
                && IsEqual(first.Y, second.Y)
                && IsEqual(first.Z, second.Z);
        }
    }

    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    public class PlaceBottomElbowData
    {
        public SourcePipeData MainPipe { get; set; }
        public List<ServicePlaceBottomElbow> ServicePlaceBottomElbows { get; set; }

        public ServicePlaceBottomElbow ServicePlaceBottomElbowFinal { get; set; }
        public SourcePipeData Pipe1_Final { get; set; }
        public SourcePipeData Pipe2_Final { get; set; }

        public SourcePipeData VerticalPipe { get; set; }
        public PipeType PipeTypeProcess { get; set; }

        private XYZ m_intsMain1_pnt = null;
        private XYZ m_sub1_pnt0 = null;
        private XYZ m_sub1_pnt1 = null;
        private XYZ m_sub1Ints_pnt = null;

        private XYZ m_intsMain2_pnt = null;
        private XYZ m_sub2_pnt0 = null;
        private XYZ m_sub2_pnt1 = null;
        private XYZ m_sub2Ints_pnt = null;

        private FamilyInstance m_bottomElbow = null;

        private Pipe m_primerPipe_1 = null;
        private Pipe m_primerPipe_2 = null;

        private FamilyInstance m_topTee = null;

        private FamilyInstance m_topReducer1 = null;
        private FamilyInstance m_topReducer2 = null;

        private FamilyInstance m_topElbow = null;

        public PlaceBottomElbowData(Pipe mainPipe, List<PairPipes> pairPipes, PipeType pipeTypeProcess)
        {
            ServicePlaceBottomElbows = new List<ServicePlaceBottomElbow>();
            PipeTypeProcess = pipeTypeProcess;
            if (mainPipe != null && pairPipes != null && pairPipes.Count > 0)
            {
                MainPipe = new SourcePipeData(mainPipe);

                foreach (PairPipes pairPipe in pairPipes)
                {
                    ServicePlaceBottomElbow servicePlaceBottomElbow = new ServicePlaceBottomElbow(mainPipe, pairPipe._Pipe1, pairPipe._Pipe2);
                    if (servicePlaceBottomElbow.IsValidPlaceBottomElbow)
                    {
                        ServicePlaceBottomElbows.Add(servicePlaceBottomElbow);
                    }
                }

                if (ServicePlaceBottomElbows.Count > 0)
                {
                    ServicePlaceBottomElbowFinal = ServicePlaceBottomElbows.OrderByDescending(item => item.DistanceValid).FirstOrDefault();
                    Pipe1_Final = ServicePlaceBottomElbows.OrderByDescending(item => item.DistanceValid).FirstOrDefault().Pipe1;
                    Pipe2_Final = ServicePlaceBottomElbows.OrderByDescending(item => item.DistanceValid).FirstOrDefault().Pipe2;
                    Initialize();
                }
            }
            PipeTypeProcess = pipeTypeProcess;
        }

        private void Initialize()
        {
            if (MainPipe != null && Pipe1_Final != null)
            {
                using (Transaction reTrans = new Transaction(Global.UIDoc.Document, "STEP"))
                {
                    try
                    {
                        reTrans.Start();

                        if (BeforeProcess() == false)
                        {
                            reTrans.RollBack();
                            return;
                        }

                        if (HandleVerticalPipe() == false)
                        {
                            reTrans.RollBack();
                            return;
                        }

                        if (HandleBottomElbowConnection() == false)
                        {
                            reTrans.RollBack();
                            return;
                        }

                        if (HandlerTopTee() == false)
                        {
                            reTrans.RollBack();
                            return;
                        }
                        reTrans.Commit();
                    }
                    catch (Exception)
                    {
                        reTrans.RollBack();
                    }
                }
            }
        }

        private bool BeforeProcess()
        {
            try
            {
                // Check has two pipe valid
                bool flagTwoPipe = Pipe2_Final.ProcessPipe != null ? true : false;

                if (ServicePlaceBottomElbowFinal.IntesectCase == IntesectCase.None)
                    return false;

                if (ServicePlaceBottomElbowFinal.IntesectCase == IntesectCase.VirtualMain_VirtualBranch || ServicePlaceBottomElbowFinal.IntesectCase == IntesectCase.VirtualMain_RealBranch)
                {
                    if (MainPipe.DirectionPipe == ProcessDirectionPipe.FirstToSecond)
                    {
                        (MainPipe.ProcessPipe.Location as LocationCurve).Curve = Line.CreateBound(ServicePlaceBottomElbowFinal.IntsectPoint, MainPipe.FirstConnector.Origin);
                    }
                    else if (MainPipe.DirectionPipe == ProcessDirectionPipe.SecondToFirst)
                    {
                        (MainPipe.ProcessPipe.Location as LocationCurve).Curve = Line.CreateBound(ServicePlaceBottomElbowFinal.IntsectPoint, MainPipe.SecondConnector.Origin);
                    }
                }

                if (HandlerDivide(MainPipe.ProcessPipe, Pipe1_Final.ProcessPipe, flagTwoPipe, out m_intsMain1_pnt, out m_sub1_pnt0, out m_sub1_pnt1, out m_sub1Ints_pnt) == true)
                {
                    // If only 1 pipe
                    if (Pipe2_Final.ProcessPipe == null)
                    {
                        (Pipe1_Final.ProcessPipe.Location as LocationCurve).Curve = Line.CreateBound(m_sub1_pnt0, m_sub1Ints_pnt);

                        Pipe2_Final.ProcessPipe = Common.Clone(Pipe1_Final.ProcessPipe) as Pipe;
                        (Pipe2_Final.ProcessPipe.Location as LocationCurve).Curve = Line.CreateBound(m_sub1Ints_pnt, m_sub1_pnt1);

                        // Set point value again
                        m_sub1_pnt1 = m_sub1Ints_pnt;
                    }

                    if (Pipe2_Final.ProcessPipe != null)
                    {
                        if (HandlerDivide(MainPipe.ProcessPipe, Pipe2_Final.ProcessPipe, true, out m_intsMain2_pnt, out m_sub2_pnt0, out m_sub2_pnt1, out m_sub2Ints_pnt) == false)
                        {
                            return false;
                        }

                        if (m_intsMain1_pnt.DistanceTo(m_intsMain2_pnt) > 0.001)
                        {
                            return false;
                        }

                        if (m_sub1Ints_pnt.DistanceTo(m_sub2Ints_pnt) > 0.001)
                        {
                            return false;
                        }
                    }
                }

                return true;
            }
            catch (Exception)
            { }
            return false;
        }

        private bool HandleVerticalPipe()
        {
            try
            {
                Pipe verticalPipe = Common.Clone(Pipe1_Final.ProcessPipe) as Pipe;
                verticalPipe.PipeType = PipeTypeProcess;
                double diameter = (double)GetParameterValueByName(MainPipe.ProcessPipe, "Diameter");
                verticalPipe.LookupParameter("Diameter").Set(diameter);
                (verticalPipe.Location as LocationCurve).Curve = Line.CreateBound(m_intsMain1_pnt, m_sub1Ints_pnt);
                VerticalPipe = new SourcePipeData(verticalPipe);

                return true;
            }
            catch (Exception)
            { }
            return false;
        }

        private bool HandleBottomElbowConnection()
        {
            try
            {
                if (MainPipe.ProcessPipe as Pipe != null && VerticalPipe.ProcessPipe as Pipe != null && m_intsMain1_pnt != null)
                {
                    m_bottomElbow = CreateElbowFitting(MainPipe.ProcessPipe as Pipe, VerticalPipe.ProcessPipe as Pipe, m_intsMain1_pnt);
                    if (m_bottomElbow != null)
                        return true;
                }
            }
            catch (Exception)
            { }
            return false;
        }

        private bool HandlerTopTee()
        {
            try
            {
                // Process with sub pipes

                Connector cntBottom_topTee = GetConnectorClosestTo(VerticalPipe.ProcessPipe, m_sub1Ints_pnt);

                //double diameter_primer = 15 * Common.mmToFT;

                double diameter_10 = 10 * Common.mmToFT;
                double diameter = (double)GetParameterValueByName(MainPipe.ProcessPipe, "Diameter");

                // If diameter vertical pipe = diameter main pipe
                if (g(diameter, Pipe1_Final.ProcessPipe.Diameter) && g(diameter, Pipe2_Final.ProcessPipe.Diameter))
                {
                    //Connect to sub
                    (Pipe1_Final.ProcessPipe.Location as LocationCurve).Curve = Line.CreateBound(m_sub1_pnt0, m_sub1_pnt1);
                    (Pipe2_Final.ProcessPipe.Location as LocationCurve).Curve = Line.CreateBound(m_sub2_pnt0, m_sub2_pnt1);

                    //Connect
                    Connector c3 = GetConnectorClosestTo(Pipe1_Final.ProcessPipe, m_sub1Ints_pnt);
                    Connector c4 = GetConnectorClosestTo(Pipe2_Final.ProcessPipe, m_sub2Ints_pnt);
                    m_topTee = CreatTee(c3, c4, cntBottom_topTee);
                    if (m_topTee == null)
                    {
                        return false;
                    }

                    return true;
                }
                // If diameter vertical pipe < diameter main pipe
                else
                {
                    Connector c3 = null;
                    // Create primer pipe 1
                    m_primerPipe_1 = null;

                    XYZ tempVector = null;

                    if (g(diameter, Pipe1_Final.ProcessPipe.Diameter) == false)
                    {
                        // Create primer pipe 1
                        m_primerPipe_1 = Common.Clone(VerticalPipe.ProcessPipe) as Pipe;
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
                        (Pipe1_Final.ProcessPipe.Location as LocationCurve).Curve = line;
                        c3 = Common.GetConnectorClosestTo(Pipe1_Final.ProcessPipe, m_sub1Ints_pnt);

                        tempVector = line.Direction;
                    }

                    Connector c4 = null;
                    // Create primer pipe 2
                    m_primerPipe_2 = null;

                    if (g(diameter, Pipe2_Final.ProcessPipe.Diameter) == false)
                    {
                        // Create primer pipe 1
                        m_primerPipe_2 = Common.Clone(VerticalPipe.ProcessPipe) as Pipe;
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
                        (Pipe2_Final.ProcessPipe.Location as LocationCurve).Curve = line;
                        c4 = Common.GetConnectorClosestTo(Pipe1_Final.ProcessPipe, m_sub1Ints_pnt);

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
                    HandlerConnectTopReducer(Pipe1_Final.ProcessPipe,
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
                    HandlerConnectTopReducer(Pipe2_Final.ProcessPipe,
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

                    return true;
                }
            }
            catch (Exception)
            { }
            return false;
        }

        /// <summary>
        /// Handler Connect Top Reducer
        /// </summary>
        /// <param name="firstPipe"></param>
        /// <param name="primerPipe_1"></param>
        /// <param name="checkPrimerTemp1"></param>
        /// <param name="checkPrimerTemp2"></param>
        /// <param name="main1_p"></param>
        /// <param name="v_m_1"></param>
        /// <param name="main2_p"></param>
        /// <param name="v_m_2"></param>
        /// <param name="diameter_10"></param>
        /// <param name="sub1Ints_pnt"></param>
        /// <param name="sub1_pnt0"></param>
        /// <param name="sub1_pnt1"></param>
        /// <param name="tee"></param>
        /// <param name="reducer_1"></param>
        /// <returns></returns>
        private bool HandlerConnectTopReducer(Pipe firstPipe,
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

        /// <summary>
        /// g
        /// </summary>
        /// <param name="d1"></param>
        /// <param name="d2"></param>
        /// <returns></returns>
        private bool g(double d1, double d2)
        {
            if (Math.Abs(d1 - d2) < 0.001)
                return true;

            return false;
        }

        /// <summary>
        /// Create Tee
        /// </summary>
        /// <param name="c3"></param>
        /// <param name="c4"></param>
        /// <param name="c5"></param>
        /// <returns></returns>
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

        /// <summary>
        /// Get pipe connect
        /// </summary>
        /// <param name="cs"></param>
        /// <returns></returns>
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

        /// <summary>
        /// Set Location Line
        /// </summary>
        /// <param name="pipe"></param>
        /// <param name="pOn"></param>
        /// <param name="pOther"></param>
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

        /// <summary>
        /// Creat Transition Fitting
        /// </summary>
        /// <param name="mep1"></param>
        /// <param name="mep2"></param>
        /// <param name="checkDistance"></param>
        /// <returns></returns>
        private FamilyInstance CreatTransitionFitting(MEPCurve mep1, MEPCurve mep2, bool checkDistance = false)
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

        /// <summary>
        /// Is Same
        /// </summary>
        /// <param name="mep1"></param>
        /// <param name="mep2"></param>
        /// <returns></returns>
        private bool iss(MEPCurve mep1, MEPCurve mep2)
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

        /// <summary>
        /// Create Elbow Fitting
        /// </summary>
        /// <param name="pipeMain"></param>
        /// <param name="pipeCurrent"></param>
        /// <param name="splitPoint"></param>
        /// <param name="main2"></param>
        /// <returns></returns>
        private FamilyInstance CreateElbowFitting(Pipe pipeMain, Pipe pipeCurrent, XYZ splitPoint)
        {
            try
            {
                var curve = (pipeMain.Location as LocationCurve).Curve;

                var p0 = curve.GetEndPoint(0);
                var p1 = curve.GetEndPoint(1);

                var pipeTempMain1 = pipeMain;

                //Connect
                var c3 = GetConnectorClosestTo(pipeTempMain1, splitPoint);

                var c5 = GetConnectorClosestTo(pipeCurrent, splitPoint);

                try
                {
                    FamilyInstance retFitting = null;
                    retFitting = Global.UIDoc.Document.Create.NewElbowFitting(c3, c5);
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
        private Connector GetConnectorClosestTo(Element e,
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
            Global.UIDoc.Document.Regenerate();
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
        private dynamic GetParameterValueByName(Element elem, string paramName)
        {
            if (elem != null)
            {
                Parameter parameter = elem.LookupParameter(paramName);
                return GetParameterValue(parameter);
            }
            return null;
        }
    }

    public class ServicePlaceBottomElbow
    {
        public SourcePipeData Pipe1 { get; set; }
        public SourcePipeData Pipe2 { get; set; }
        public SourcePipeData MainPipe { get; set; }
        public bool IsValidPlaceBottomElbow { get; set; }
        public IntesectCase IntesectCase { get; set; }
        public XYZ IntsectPoint { get; set; }
        public XYZ IntsectPoint_2d { get; set; }
        public double DistanceValid { get; set; }

        public ServicePlaceBottomElbow(Pipe mainPipe, Pipe pipe1, Pipe pipe2)
        {
            IsValidPlaceBottomElbow = false;
            DistanceValid = double.MinValue;

            if (mainPipe != null && pipe1 != null)
            {
                MainPipe = new SourcePipeData(mainPipe);
                Pipe1 = new SourcePipeData(pipe1);
                Pipe2 = new SourcePipeData(pipe2);
                XYZ intersectPnt_2d = null;
                XYZ intersectPnt = null;

                if (MainPipe.DirectionPipe == ProcessDirectionPipe.FirstToSecond || MainPipe.DirectionPipe == ProcessDirectionPipe.SecondToFirst)

                {
                    if (RealityIntersect(MainPipe.CurvePipe_2d, Pipe1.CurvePipe_2d, out intersectPnt_2d))
                    {
                        IsValidPlaceBottomElbow = true;
                        IntesectCase = IntesectCase.RealIntersect;
                    }
                    else if (RealityIntersect(MainPipe.CurvePipe_2d, Pipe1.CurvePipe_Unbound_2d, out intersectPnt_2d))
                    {
                        IsValidPlaceBottomElbow = true;
                        IntesectCase = IntesectCase.RealMain_VirtualBranch;
                    }
                    else if (RealityIntersect(MainPipe.CurvePipe_Unbound_2d, Pipe1.CurvePipe_2d, out intersectPnt_2d))
                    {
                        IsValidPlaceBottomElbow = true;
                        IntesectCase = IntesectCase.VirtualMain_RealBranch;
                    }
                    else if (RealityIntersect(MainPipe.CurvePipe_Unbound_2d, Pipe1.CurvePipe_Unbound_2d_Limit, out intersectPnt_2d))
                    {
                        IsValidPlaceBottomElbow = true;
                        IntesectCase = IntesectCase.VirtualMain_VirtualBranch;
                    }

                    IntsectPoint_2d = intersectPnt_2d;

                    if (IsValidPlaceBottomElbow = true && IntsectPoint_2d != null)
                    {
                        double temp = 200;

                        var lineZ = Line.CreateBound(new XYZ(IntsectPoint_2d.X, IntsectPoint_2d.Y, IntsectPoint_2d.Z - temp), new XYZ(IntsectPoint_2d.X, IntsectPoint_2d.Y, IntsectPoint_2d.Z + temp));

                        if (RealityIntersect(MainPipe.CurvePipe_Unbound, lineZ, out intersectPnt))
                        {
                            IntsectPoint = intersectPnt;
                        }

                        if (MainPipe.DirectionPipe == ProcessDirectionPipe.FirstToSecond)
                        {
                            DistanceValid = IntsectPoint_2d.DistanceTo(ToPoint2D(MainPipe.FirstConnector.Origin));
                        }
                        else if (MainPipe.DirectionPipe == ProcessDirectionPipe.SecondToFirst)
                        {
                            DistanceValid = IntsectPoint_2d.DistanceTo(ToPoint2D(MainPipe.SecondConnector.Origin));
                        }
                    }
                }
            }
        }

        private bool RealityIntersect(Line mainLine, Line checkLine, out XYZ intersectPoint)
        {
            intersectPoint = null;
            try
            {
                IntersectionResultArray intsRetArr = new IntersectionResultArray();
                var intsRet = mainLine.Intersect(checkLine, out intsRetArr);
                if (intsRet == SetComparisonResult.Overlap)
                {
                    intersectPoint = intsRetArr.get_Item(0).XYZPoint;
                    return true;
                }
            }
            catch (Exception)
            { }
            return false;
        }

        private XYZ ToPoint2D(XYZ point3d, double z = 0)
        {
            return new XYZ(point3d.X, point3d.Y, z);
        }
    }

    public enum IntesectCase : int
    {
        None = -1,
        RealIntersect,
        RealMain_VirtualBranch,
        VirtualMain_RealBranch,
        VirtualMain_VirtualBranch
    }

    public class SourcePipeData
    {
        public double mmToFT = 0.0032808399;
        public Pipe ProcessPipe { get; set; }
        public Connector FirstConnector { get; set; }
        public Connector SecondConnector { get; set; }
        public ProcessDirectionPipe DirectionPipe { get; set; }
        public Line CurvePipe_Reality { get; set; }
        public Line CurvePipe_Unbound { get; set; }
        public Line CurvePipe_Unbound_Limit { get; set; }
        public Line CurvePipe_2d { get; set; }
        public Line CurvePipe_Unbound_2d { get; set; }
        public Line CurvePipe_Unbound_2d_Limit { get; set; }

        public SourcePipeData(Pipe processPipe)
        {
            DirectionPipe = DirectionPipe = ProcessDirectionPipe.None;
            if (processPipe != null)
            {
                ProcessPipe = processPipe;
                Initialize();
            }
        }

        private void Initialize()
        {
            try
            {
                if (ProcessPipe != null && ProcessPipe.ConnectorManager != null)
                {
                    List<Connector> cntOfPipe = GetConnectors(ProcessPipe.ConnectorManager.Connectors);
                    if (cntOfPipe.Count >= 2)
                    {
                        FirstConnector = cntOfPipe[0];
                        SecondConnector = cntOfPipe[1];

                        if (FirstConnector.IsConnected == true && SecondConnector.IsConnected == false)
                        {
                            DirectionPipe = ProcessDirectionPipe.FirstToSecond;
                        }
                        else if (FirstConnector.IsConnected == false && SecondConnector.IsConnected == true)
                        {
                            DirectionPipe = ProcessDirectionPipe.SecondToFirst;
                        }
                        else if (FirstConnector.IsConnected == false && SecondConnector.IsConnected == false)
                        {
                            DirectionPipe = ProcessDirectionPipe.BothExpand;
                        }
                        else
                        {
                            DirectionPipe = ProcessDirectionPipe.None;
                        }

                        if (DirectionPipe != ProcessDirectionPipe.None)
                        {
                            XYZ direction = null;
                            XYZ direction_2d = null;
                            CurvePipe_Reality = Line.CreateBound(FirstConnector.Origin, SecondConnector.Origin);
                            CurvePipe_2d = Line.CreateBound(ToPoint2D(FirstConnector.Origin), ToPoint2D(SecondConnector.Origin));
                            direction = (SecondConnector.Origin - FirstConnector.Origin).Normalize();
                            direction_2d = (ToPoint2D(SecondConnector.Origin) - ToPoint2D(FirstConnector.Origin)).Normalize();

                            if (DirectionPipe == ProcessDirectionPipe.FirstToSecond)
                            {
                                CurvePipe_Unbound = Line.CreateBound(FirstConnector.Origin, SecondConnector.Origin + direction * 100);
                                CurvePipe_Unbound_Limit = Line.CreateBound(FirstConnector.Origin, SecondConnector.Origin + direction * 450 * mmToFT);
                                CurvePipe_Unbound_2d = Line.CreateBound(ToPoint2D(FirstConnector.Origin), ToPoint2D(SecondConnector.Origin) + direction_2d * 100);
                                CurvePipe_Unbound_2d_Limit = Line.CreateBound(ToPoint2D(FirstConnector.Origin), ToPoint2D(SecondConnector.Origin) + direction_2d * 450 * mmToFT);
                            }
                            else if (DirectionPipe == ProcessDirectionPipe.SecondToFirst)
                            {
                                CurvePipe_Unbound = Line.CreateBound(FirstConnector.Origin - direction * 100, SecondConnector.Origin);
                                CurvePipe_Unbound_Limit = Line.CreateBound(FirstConnector.Origin - direction * 450 * mmToFT, SecondConnector.Origin);
                                CurvePipe_Unbound_2d = Line.CreateBound(ToPoint2D(FirstConnector.Origin) - direction_2d * 100, ToPoint2D(SecondConnector.Origin));
                                CurvePipe_Unbound_2d_Limit = Line.CreateBound(ToPoint2D(FirstConnector.Origin) - direction_2d * 450 * mmToFT, ToPoint2D(SecondConnector.Origin));
                            }
                        }
                    }
                }
            }
            catch (Exception) { }
        }

        private XYZ ToPoint2D(XYZ point3d, double z = 0)
        {
            return new XYZ(point3d.X, point3d.Y, z);
        }

        private List<Connector> GetConnectors(ConnectorSet connectorSet)
        {
            try
            {
                List<Connector> retVal = new List<Connector>();
                foreach (Connector connector in connectorSet)
                {
                    if (connector.ConnectorType != ConnectorType.End)
                        continue;
                    retVal.Add(connector);
                }
                return retVal;
            }
            catch (Exception)
            { }
            return new List<Connector>();
        }
    }

    public enum ProcessDirectionPipe : int
    {
        None = -1,
        FirstToSecond,
        SecondToFirst,
        BothExpand
    }

    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
}