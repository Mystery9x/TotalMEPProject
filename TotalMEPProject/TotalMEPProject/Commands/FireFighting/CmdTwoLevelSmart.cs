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
using TotalMEPProject.Commands.TotalMEP;
using TotalMEPProject.Services;
using TotalMEPProject.Ultis;

namespace TotalMEPProject.Commands.FireFighting
{
    [Transaction(TransactionMode.Manual)]
    public class CmdTwoLevelSmart : IExternalCommand
    {
        private static List<ElementId> m_mainPipeIds = new List<ElementId>();

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
        private static List<PairPipes> PairPipe(List<Pipe> pipes)
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

                var expand = inforPipe1.CurveSourcePipe_Extend_Round;

                foreach (InforPie inforPipe2 in inforPipes)
                {
                    if (inforPipe1.SourcePipe.Id == inforPipe2.SourcePipe.Id)
                        continue;

                    var curve_sub2 = inforPipe2.CurveSourcePipe_Extend_Round;
                    IntersectionResultArray arr = new IntersectionResultArray();
                    var result = curve_sub2.Intersect(expand, out arr);

                    if (result != SetComparisonResult.Disjoint)
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

                m_mainPipeIds.Clear();

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
                List<DuoBranchPipe> pairPies_1 = PairingPipesPlus(branchPipes);
                if (pairPies_1 == null || pairPies_1.Count() <= 0)
                    return Result.Cancelled;

                List<PairPipes> pairPies = PairPipe(branchPipes);
                if (pairPies == null || pairPies.Count() <= 0)
                    return Result.Cancelled;

                TransactionGroup reTransGrp = new TransactionGroup(Global.UIDoc.Document, "TWO_LEVEL_SMART_COMMAND");
                try
                {
                    reTransGrp.Start();

                    if (App.m_2LevelSmartForm.DialogResultData.IsCheckedEleDiff)
                    {
                        if (GetPreferredJunctionType(mainPipes[0]) == PreferredJunctionType.Tap && App.m_2LevelSmartForm.DialogResultData.IsCheckedTeeTap)
                        {
                            ProcessWithTap(inforMainPipes, pairPies_1, App.m_2LevelSmartForm.DialogResultData);
                        }
                        else if (GetPreferredJunctionType(mainPipes[0]) == PreferredJunctionType.Tee && App.m_2LevelSmartForm.DialogResultData.IsCheckedTeeTap)
                        {
                            ProcessWithTee(inforMainPipes, pairPies_1, App.m_2LevelSmartForm.DialogResultData);
                        }
                        //else if (GetPreferredJunctionType(mainPipes[0]) == PreferredJunctionType.Tap && App.m_2LevelSmartForm.DialogResultData.IsCheckedTeeTap)
                        //{
                        //}
                        else
                        {
                            // Get ElementId main pipes
                            List<ElementId> mainPipeIds = mainPipes.Where(item => item.Id != ElementId.InvalidElementId).Select(item => item.Id).ToList();
                            m_mainPipeIds.AddRange(mainPipeIds);

                            HandleProcessTwoLevelSmartCommand handleProcess = new HandleProcessTwoLevelSmartCommand(m_mainPipeIds, pairPies, App.m_2LevelSmartForm.DialogResultData);
                        }
                    }
                    else
                    {
                        List<ElementId> mainPipeIds = mainPipes.Where(item => item.Id != ElementId.InvalidElementId).Select(item => item.Id).ToList();
                        m_mainPipeIds.AddRange(mainPipeIds);

                        HandleProcessTwoLevelSmartCommand handleProcess = new HandleProcessTwoLevelSmartCommand(m_mainPipeIds, pairPies, App.m_2LevelSmartForm.DialogResultData);
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

        private static bool ProcessSameElevation(List<InforPie> inforMainPipes, List<DuoBranchPipe> pairPies, TwoLevelSmartDialogData dialogResultData)
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
                        try
                        {
                            if (duoBranchs.SecondPipe != null)
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
                            else
                            {
                                if (RealityIntersect(flatten_mainCurve, duoBranchs.FlattenValidCurve, out XYZ flattenIntPnt1))
                                {
                                    DuoBranchPipe duoBranchPipe = new DuoBranchPipe(duoBranchs.FirstPipe, duoBranchs.SecondPipe)
                                    {
                                        FlattenIntersectPoint = flattenIntPnt1
                                    };

                                    sourceMain.Branches.Add(duoBranchPipe);
                                    sourceMain.Branches_Special.Add(duoBranchPipe);
                                }
                            }
                        }
                        catch (Exception)
                        {
                        }
                        continue;

                        //using (Transaction reTrans = new Transaction(Global.UIDoc.Document, "ZZZ"))
                        //{
                        //    reTrans.Start();
                        //    Global.UIDoc.Document.Create.NewDetailCurve(Global.UIDoc.ActiveView, duoBranchs.FlattenValidCurve);
                        //    Global.UIDoc.Document.Create.NewDetailCurve(Global.UIDoc.ActiveView, flatten_mainCurve_Extend);
                        //    reTrans.Commit();
                        //}
                    }
                    if (dialogResultData.FlagAddElbowLastBranch)
                        sourceMain.Initialize();
                    sourceMainPipes.Add(sourceMain);
                }

                foreach (var sourceMain in sourceMainPipes)
                {
                    using (Transaction reTrans = new Transaction(Global.UIDoc.Document, "STEP_1_TWO_LEVEL_SMART"))
                    {
                        reTrans.Start();
                        FailureHandlingOptions fhOpts = reTrans.GetFailureHandlingOptions();
                        List<DuoBranchPipe> allVerticalPipes = new List<DuoBranchPipe>();
                        List<DuoBranchPipe> norVerticalPipes = new List<DuoBranchPipe>();

                        // Create elbow of 2 end of the line
                        if (sourceMain.IntDuoBranch_1 != null)
                        {
                            if (sourceMain.IntDuoBranch_1.FirstPipe.SourcePipe != null)
                            {
                            }
                        }

                        if (sourceMain.IntDuoBranch_2 != null)
                        {
                        }

                        // Create vertical pipe
                        foreach (var duoBrandPipe in sourceMain.Branches_Special)
                        {
                            Pipe verticalPipe = CreateVerticalPipeWithTap(sourceMain.MainPipe, duoBrandPipe, dialogResultData.PipeSize * Common.mmToFT);
                            if (verticalPipe != null)
                            {
                                DuoBranchPipe duoBranchPipe = duoBrandPipe.Clone() as DuoBranchPipe;
                                duoBranchPipe.VerticalPipe = new InforPie(verticalPipe);
                                allVerticalPipes.Add(duoBranchPipe);
                                norVerticalPipes.Add(duoBranchPipe);
                            }
                        }
                        Global.UIDoc.Document.Regenerate();

                        foreach (var verPipe in norVerticalPipes)
                        {
                            if (!CreateTap(sourceMain.MainPipe.SourcePipe as MEPCurve, verPipe.VerticalPipe.SourcePipe as MEPCurve))
                            {
                                allVerticalPipes.Remove(allVerticalPipes.Where(item => item.VerticalPipe.SourcePipe.Id == verPipe.VerticalPipe.SourcePipe.Id).FirstOrDefault());
                                Global.UIDoc.Document.Delete(verPipe.VerticalPipe.SourcePipe.Id);
                            }
                        }

                        Global.UIDoc.Document.Regenerate();

                        //Create Top Tee
                        foreach (var duoBranch in allVerticalPipes)
                        {
                            CreateTopTee(duoBranch, dialogResultData);
                        }

                        //System.Windows.Forms.MessageBox.Show(sourceMain.MainPipe.SourcePipe.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsValueString());

                        GetInfoWarning supWarning = new GetInfoWarning(true);
                        fhOpts.SetFailuresPreprocessor(supWarning);
                        reTrans.SetFailureHandlingOptions(fhOpts);
                        reTrans.Commit();
                    }
                }
            }
            catch (Exception)
            { }
            return false;
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

        private static bool ProcessWithTee(List<InforPie> inforMainPipes, List<DuoBranchPipe> pairPies, TwoLevelSmartDialogData dialogResultData)
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
                        try
                        {
                            if (duoBranchs.SecondPipe != null)
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
                            else
                            {
                                if (RealityIntersect(flatten_mainCurve, duoBranchs.FlattenValidCurve, out XYZ flattenIntPnt1))
                                {
                                    DuoBranchPipe duoBranchPipe = new DuoBranchPipe(duoBranchs.FirstPipe, duoBranchs.SecondPipe)
                                    {
                                        FlattenIntersectPoint = flattenIntPnt1
                                    };

                                    sourceMain.Branches.Add(duoBranchPipe);
                                    sourceMain.Branches_Special.Add(duoBranchPipe);
                                }
                                else if (RealityIntersect(flatten_mainCurve_Extend, duoBranchs.FlattenValidCurve, out XYZ flattenIntPnt2))
                                {
                                    DuoBranchPipe duoBranchPipe = new DuoBranchPipe(duoBranchs.FirstPipe, duoBranchs.SecondPipe)
                                    {
                                        FlattenIntersectPoint = flattenIntPnt2
                                    };
                                    sourceMain.Branches_Special.Add(duoBranchPipe);
                                }
                            }
                        }
                        catch (Exception)
                        {
                        }
                        continue;
                    }
                    if (dialogResultData.FlagAddElbowLastBranch)
                        sourceMain.Initialize();
                    sourceMainPipes.Add(sourceMain);
                }

                double diamterPipe = dialogResultData.PipeSize * Common.mmToFT;
                PipeType pipeType = Global.UIDoc.Document.GetElement(dialogResultData.PipeTypeId) as PipeType;
                foreach (var sourceMain in sourceMainPipes)
                {
                    using (Transaction reTrans = new Transaction(Global.UIDoc.Document, "STEP_1_TWO_LEVEL_SMART"))
                    {
                        reTrans.Start();
                        FailureHandlingOptions fhOpts = reTrans.GetFailureHandlingOptions();
                        List<DuoBranchPipe> allVerticalPipes = new List<DuoBranchPipe>();
                        List<DuoBranchPipe> norVerticalPipes = new List<DuoBranchPipe>();
                        if (GetPreferredJunctionType(sourceMain.MainPipe.SourcePipe) == PreferredJunctionType.Tee)
                        {
                            // Create vertical Pipe
                            if (dialogResultData.FlagAddElbowLastBranch)
                            {
                                // Create elbow of 2 end of the line
                                if (sourceMain.IntDuoBranch_1 != null)
                                {
                                    Pipe verticalPipe_1 = CreateVerticalPipeWithElbowLastBranch(sourceMain.MainPipe, sourceMain.IntDuoBranch_1, dialogResultData.PipeSize * Common.mmToFT);
                                    if (verticalPipe_1 != null)
                                    {
                                        DuoBranchPipe duoBranchPipe = sourceMain.IntDuoBranch_1.Clone() as DuoBranchPipe;
                                        duoBranchPipe.VerticalPipe = new InforPie(verticalPipe_1);
                                        allVerticalPipes.Add(duoBranchPipe);
                                    }
                                }

                                if (sourceMain.IntDuoBranch_2 != null)
                                {
                                    Pipe verticalPipe_2 = CreateVerticalPipeWithElbowLastBranch(sourceMain.MainPipe, sourceMain.IntDuoBranch_2, dialogResultData.PipeSize * Common.mmToFT);
                                    if (verticalPipe_2 != null)
                                    {
                                        DuoBranchPipe duoBranchPipe = sourceMain.IntDuoBranch_2.Clone() as DuoBranchPipe;
                                        duoBranchPipe.VerticalPipe = new InforPie(verticalPipe_2);
                                        allVerticalPipes.Add(duoBranchPipe);
                                    }
                                }
                            }

                            //Create Top Tee Last Branch
                            foreach (var duoBranch in allVerticalPipes)
                            {
                                CreateTopTee(duoBranch, dialogResultData);
                            }

                            List<ElementId> mainPipeIds = new List<ElementId>();
                            mainPipeIds.Add(sourceMain.MainPipe.SourcePipe.Id);

                            // Create vertical pipe
                            foreach (var duoBrandPipe in sourceMain.Branches_Special)
                            {
                                Pipe pipe1 = duoBrandPipe.FirstPipe.SourcePipe;
                                Pipe pipe2 = duoBrandPipe.SecondPipe != null ? duoBrandPipe.SecondPipe.SourcePipe : null;

                                using (SubTransaction reSubTrans = new SubTransaction(Global.UIDoc.Document))
                                {
                                    reSubTrans.Start();
                                    try
                                    {
                                        List<Pipe> mainPipes = new List<Pipe>();
                                        mainPipeIds.ForEach(item =>
                                        {
                                            mainPipes.Add(Global.UIDoc.Document.GetElement(item) as Pipe);
                                        });

                                        bool twoPipes = pipe2 != null ? true : false;
                                        Pipe validMainPipe = ProcessMainPipe(mainPipes, pipe1, out bool flagSplit, out XYZ intPnt);

                                        XYZ inter_main1 = null;
                                        XYZ p0_sub1 = null;
                                        XYZ p1_sub1 = null;
                                        XYZ inters_sub1 = null;
                                        Pipe main = null;

                                        foreach (ElementId pId in mainPipeIds)
                                        {
                                            var m = Global.UIDoc.Document.GetElement(pId) as Pipe;
                                            if (DivideCase_2(m, pipe1, twoPipes, out inter_main1, out p0_sub1, out p1_sub1, out inters_sub1) == true)
                                            {
                                                main = m;
                                                break;
                                            }
                                        }

                                        if (pipe2 == null)
                                        {
                                            var temp_pipe1 = pipe1;
                                            var temp_pipe2_id = PlumbingUtils.BreakCurve(Global.UIDoc.Document, temp_pipe1.Id, inters_sub1);
                                            var temp_pipe2 = Global.UIDoc.Document.GetElement(temp_pipe2_id) as Pipe;
                                            pipe1 = temp_pipe1;
                                            pipe2 = temp_pipe2;
                                        }

                                        DivideCase_2(main, pipe1, twoPipes, out inter_main1, out p0_sub1, out p1_sub1, out inters_sub1);

                                        XYZ inter_main2 = null;
                                        XYZ p0_sub2 = null;
                                        XYZ p1_sub2 = null;
                                        XYZ inters_sub2 = null;

                                        if (DivideCase_2(main, pipe2, true, out inter_main2, out p0_sub2, out p1_sub2, out inters_sub2) == false)
                                        {
                                            reSubTrans.RollBack();
                                            continue;
                                        }

                                        if (inter_main1.DistanceTo(inter_main2) > 0.001)
                                        {
                                            reSubTrans.RollBack();
                                            continue; //Check
                                        }

                                        if (inters_sub1.DistanceTo(inters_sub2) > 0.001)
                                        {
                                            reSubTrans.RollBack();
                                            continue; //Check
                                        }

                                        //Create vertical pipe
                                        var vertical = Common.Clone(pipe1) as Pipe;
                                        vertical.PipeType = pipeType;
                                        vertical.LookupParameter("Diameter").Set(diamterPipe);
                                        (vertical.Location as LocationCurve).Curve = Line.CreateBound(inter_main1, inters_sub1);

                                        //Conect to main pipe

                                        //Try
                                        if (main as Pipe != null && vertical as Pipe != null && inter_main1 != null)
                                        {
                                            Pipe main2 = null;

                                            var fitting = ee(main as Pipe, vertical as Pipe, inter_main1, out main2);
                                            if (main2 != null && fitting != null)
                                            {
                                                mainPipeIds.Add(main2.Id);
                                            }
                                            else
                                            {
                                                reSubTrans.RollBack();
                                                continue;
                                            }
                                        }

                                        //Process for sub pipes
                                        var c5 = Common.GetConnectorClosestTo(vertical, inters_sub1);

                                        double dMoi = 15 * Common.mmToFT;

                                        double d10 = 10 * Common.mmToFT;
                                        double dtemp = 2;

                                        if (CompareDouble(diamterPipe, pipe1.Diameter) && CompareDouble(diamterPipe, pipe2.Diameter))
                                        {
                                            //Connect to sub
                                            (pipe1.Location as LocationCurve).Curve = Line.CreateBound(p0_sub1, p1_sub1);
                                            (pipe2.Location as LocationCurve).Curve = Line.CreateBound(p0_sub2, p1_sub2);

                                            //Connect
                                            var c3 = Common.GetConnectorClosestTo(pipe1, inters_sub1);
                                            var c4 = Common.GetConnectorClosestTo(pipe2, inters_sub1);

                                            if (CreateFittingForMEPUtils.ct(c3, c4, c5) == null)
                                            {
                                                reSubTrans.RollBack();
                                                continue;
                                            }
                                        }
                                        else
                                        {
                                            Connector c3 = null;
                                            Pipe pipe_moi_1 = null;

                                            XYZ v = null;
                                            if (CompareDouble(diamterPipe, pipe1.Diameter) == false)
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

                                            if (CompareDouble(diamterPipe, pipe2.Diameter) == false)
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
                                            FamilyInstance fitting = CreateFittingForMEPUtils.ct(c3, c4, c5);
                                            if (fitting == null)
                                            {
                                                reSubTrans.RollBack();
                                                continue;
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

                                                FamilyInstance topReducer_1 = null;
                                                FamilyInstance topReducer_2 = null;
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

                                                        topReducer_1 = CreateFittingForMEPUtils.ctt(pipe1, pipe_moi_1);
                                                    }
                                                    else if (mep2 != null && pipe_moi_1.Id == mep2.Id)
                                                    {
                                                        var line = Line.CreateUnbound(main2_p, v_m_2 * 10);

                                                        var p = line.Evaluate(d10, false);
                                                        (pipe_moi_1.Location as LocationCurve).Curve = Line.CreateBound(main2_p, p);

                                                        var futher = inters_sub1.DistanceTo(p0_sub1) > inters_sub1.DistanceTo(p1_sub1) ? p0_sub1 : p1_sub1;

                                                        ll(pipe1, futher, p);

                                                        topReducer_1 = CreateFittingForMEPUtils.ctt(pipe1, pipe_moi_1);
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

                                                        topReducer_2 = CreateFittingForMEPUtils.ctt(pipe2, pipe_moi_2);
                                                    }
                                                    else if (mep2 != null && pipe_moi_2.Id == mep2.Id)
                                                    {
                                                        var line = Line.CreateUnbound(main2_p, v_m_2 * 10);

                                                        var p = line.Evaluate(d10, false);
                                                        (pipe_moi_2.Location as LocationCurve).Curve = Line.CreateBound(main2_p, p);

                                                        var futher = inters_sub2.DistanceTo(p0_sub2) > inters_sub2.DistanceTo(p1_sub2) ? p0_sub2 : p1_sub2;

                                                        ll(pipe2, futher, p);

                                                        topReducer_2 = CreateFittingForMEPUtils.ctt(pipe2, pipe_moi_2);
                                                    }
                                                }

                                                if (dialogResultData.FlagAddNipple == true && dialogResultData.NippleFamily != null)
                                                {
                                                    // Nipple Family
                                                    HandlerConnectAccessoryNipple(pipe_moi_1, topReducer_1, fitting, dialogResultData.NippleFamily);
                                                    HandlerConnectAccessoryNipple(pipe_moi_2, topReducer_2, fitting, dialogResultData.NippleFamily);
                                                }
                                            }
                                        }
                                        reSubTrans.Commit();
                                    }
                                    catch (Exception)
                                    {
                                        reSubTrans.RollBack();
                                    }
                                }
                            }
                            Global.UIDoc.Document.Regenerate();
                        }
                        GetInfoWarning supWarning = new GetInfoWarning(true);
                        fhOpts.SetFailuresPreprocessor(supWarning);
                        reTrans.SetFailureHandlingOptions(fhOpts);
                        reTrans.Commit();
                    }
                }
            }
            catch (Exception)
            { }
            return false;
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
                        try
                        {
                            if (duoBranchs.SecondPipe != null)
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
                            else
                            {
                                if (RealityIntersect(flatten_mainCurve, duoBranchs.FlattenValidCurve, out XYZ flattenIntPnt1))
                                {
                                    DuoBranchPipe duoBranchPipe = new DuoBranchPipe(duoBranchs.FirstPipe, duoBranchs.SecondPipe)
                                    {
                                        FlattenIntersectPoint = flattenIntPnt1
                                    };

                                    sourceMain.Branches.Add(duoBranchPipe);
                                    sourceMain.Branches_Special.Add(duoBranchPipe);
                                }
                            }
                        }
                        catch (Exception)
                        {
                        }
                        continue;

                        //using (Transaction reTrans = new Transaction(Global.UIDoc.Document, "ZZZ"))
                        //{
                        //    reTrans.Start();
                        //    Global.UIDoc.Document.Create.NewDetailCurve(Global.UIDoc.ActiveView, duoBranchs.FlattenValidCurve);
                        //    Global.UIDoc.Document.Create.NewDetailCurve(Global.UIDoc.ActiveView, flatten_mainCurve_Extend);
                        //    reTrans.Commit();
                        //}
                    }
                    if (dialogResultData.FlagAddElbowLastBranch)
                        sourceMain.Initialize();
                    sourceMainPipes.Add(sourceMain);
                }

                foreach (var sourceMain in sourceMainPipes)
                {
                    using (Transaction reTrans = new Transaction(Global.UIDoc.Document, "STEP_1_TWO_LEVEL_SMART"))
                    {
                        reTrans.Start();
                        FailureHandlingOptions fhOpts = reTrans.GetFailureHandlingOptions();
                        List<DuoBranchPipe> allVerticalPipes = new List<DuoBranchPipe>();
                        List<DuoBranchPipe> norVerticalPipes = new List<DuoBranchPipe>();
                        if (GetPreferredJunctionType(sourceMain.MainPipe.SourcePipe) == PreferredJunctionType.Tap)
                        {
                            // Create vertical Pipe
                            if (dialogResultData.FlagAddElbowLastBranch)
                            {
                                // Create elbow of 2 end of the line
                                if (sourceMain.IntDuoBranch_1 != null)
                                {
                                    Pipe verticalPipe_1 = CreateVerticalPipeWithElbowLastBranch(sourceMain.MainPipe, sourceMain.IntDuoBranch_1, dialogResultData.PipeSize * Common.mmToFT);
                                    if (verticalPipe_1 != null)
                                    {
                                        DuoBranchPipe duoBranchPipe = sourceMain.IntDuoBranch_1.Clone() as DuoBranchPipe;
                                        duoBranchPipe.VerticalPipe = new InforPie(verticalPipe_1);
                                        allVerticalPipes.Add(duoBranchPipe);
                                    }
                                }

                                if (sourceMain.IntDuoBranch_2 != null)
                                {
                                    Pipe verticalPipe_2 = CreateVerticalPipeWithElbowLastBranch(sourceMain.MainPipe, sourceMain.IntDuoBranch_2, dialogResultData.PipeSize * Common.mmToFT);
                                    if (verticalPipe_2 != null)
                                    {
                                        DuoBranchPipe duoBranchPipe = sourceMain.IntDuoBranch_2.Clone() as DuoBranchPipe;
                                        duoBranchPipe.VerticalPipe = new InforPie(verticalPipe_2);
                                        allVerticalPipes.Add(duoBranchPipe);
                                    }
                                }
                            }

                            // Create vertical pipe
                            foreach (var duoBrandPipe in sourceMain.Branches_Special)
                            {
                                Pipe verticalPipe = CreateVerticalPipeWithTap(sourceMain.MainPipe, duoBrandPipe, dialogResultData.PipeSize * Common.mmToFT);
                                if (verticalPipe != null)
                                {
                                    DuoBranchPipe duoBranchPipe = duoBrandPipe.Clone() as DuoBranchPipe;
                                    duoBranchPipe.VerticalPipe = new InforPie(verticalPipe);
                                    allVerticalPipes.Add(duoBranchPipe);
                                    norVerticalPipes.Add(duoBranchPipe);
                                }
                            }
                            Global.UIDoc.Document.Regenerate();

                            foreach (var verPipe in norVerticalPipes)
                            {
                                if (!CreateTap(sourceMain.MainPipe.SourcePipe as MEPCurve, verPipe.VerticalPipe.SourcePipe as MEPCurve))
                                {
                                    allVerticalPipes.Remove(allVerticalPipes.Where(item => item.VerticalPipe.SourcePipe.Id == verPipe.VerticalPipe.SourcePipe.Id).FirstOrDefault());
                                    Global.UIDoc.Document.Delete(verPipe.VerticalPipe.SourcePipe.Id);
                                }
                            }

                            Global.UIDoc.Document.Regenerate();

                            //Create Top Tee
                            foreach (var duoBranch in allVerticalPipes)
                            {
                                CreateTopTee(duoBranch, dialogResultData);
                            }

                            //System.Windows.Forms.MessageBox.Show(sourceMain.MainPipe.SourcePipe.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsValueString());
                        }
                        GetInfoWarning supWarning = new GetInfoWarning(true);
                        fhOpts.SetFailuresPreprocessor(supWarning);
                        reTrans.SetFailureHandlingOptions(fhOpts);
                        reTrans.Commit();
                    }
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

        private static Pipe CreateVerticalPipeWithElbowLastBranch(InforPie mainPipe, DuoBranchPipe duoBranchPipe, double diameter)
        {
            try
            {
                if (mainPipe == null || duoBranchPipe == null)
                    return null;

                Pipe retVal = null;
                double temp = 200;

                var lineZ = Line.CreateBound(new XYZ(duoBranchPipe.FlattenIntersectPoint.X, duoBranchPipe.FlattenIntersectPoint.Y, duoBranchPipe.FlattenIntersectPoint.Z - temp)
                    , new XYZ(duoBranchPipe.FlattenIntersectPoint.X, duoBranchPipe.FlattenIntersectPoint.Y, duoBranchPipe.FlattenIntersectPoint.Z + temp));

                IntersectionResultArray intArr = new IntersectionResultArray();
                var setComRet = mainPipe.CurveSourcePipe_Unbound.Intersect(lineZ, out intArr);
                if (setComRet != SetComparisonResult.Overlap)
                    return null;
                var intMainPipe = intArr.get_Item(0).XYZPoint;

                intArr = new IntersectionResultArray();
                setComRet = duoBranchPipe.FirstPipe.CurveSourcePipe_Unbound.Intersect(lineZ, out intArr);
                if (setComRet != SetComparisonResult.Overlap)
                    return null;
                var intBranchPipe = intArr.get_Item(0).XYZPoint;

                using (SubTransaction reSubTrans = new SubTransaction(Global.UIDoc.Document))
                {
                    reSubTrans.Start();
                    var verticalPipe = Common.Clone(mainPipe.SourcePipe) as Pipe;
                    verticalPipe.LookupParameter("Diameter").Set(diameter);
                    (verticalPipe.Location as LocationCurve).Curve = Line.CreateBound(intMainPipe, intBranchPipe);

                    try
                    {
                        Connector cnt_1 = Common.GetConnectorClosestTo(verticalPipe, intMainPipe);

                        Connector cnt_2 = Common.GetConnectorClosestTo(mainPipe.SourcePipe, intMainPipe);
                        if (cnt_2.IsConnected)
                        {
                            cnt_2 = Common.GetConnectorNotConnnected(mainPipe.SourcePipe.ConnectorManager);
                        }

                        if (cnt_1 != null && !cnt_1.IsConnected && cnt_2 != null && !cnt_2.IsConnected)
                        {
                            var elbow = Global.UIDoc.Document.Create.NewElbowFitting(cnt_1, cnt_2);
                            retVal = verticalPipe;
                            reSubTrans.Commit();
                        }
                        else
                        {
                            reSubTrans.RollBack();
                        }
                    }
                    catch (Exception)
                    {
                        reSubTrans.RollBack();
                    }
                }

                return retVal;
            }
            catch (Exception)
            { }
            return null;
        }

        private static Pipe CreateVerticalPipeWithTap(InforPie mainPipe, DuoBranchPipe duoBranchPipe, double diameter)
        {
            try
            {
                if (mainPipe == null || duoBranchPipe == null)
                    return null;

                Pipe retVal = null;
                double temp = 200;

                var lineZ = Line.CreateBound(new XYZ(duoBranchPipe.FlattenIntersectPoint.X, duoBranchPipe.FlattenIntersectPoint.Y, duoBranchPipe.FlattenIntersectPoint.Z - temp)
                    , new XYZ(duoBranchPipe.FlattenIntersectPoint.X, duoBranchPipe.FlattenIntersectPoint.Y, duoBranchPipe.FlattenIntersectPoint.Z + temp));

                IntersectionResultArray intArr = new IntersectionResultArray();
                var setComRet = mainPipe.CurveSourcePipe_Unbound.Intersect(lineZ, out intArr);
                if (setComRet != SetComparisonResult.Overlap)
                    return null;
                var intMainPipe = intArr.get_Item(0).XYZPoint;

                intArr = new IntersectionResultArray();
                Line lineUn = null;
                if (duoBranchPipe.SecondPipe.CurveSourcePipe == null)
                    lineUn = Line.CreateUnbound(duoBranchPipe.FirstPipe.CurveSourcePipe.Origin, duoBranchPipe.FirstPipe.CurveSourcePipe.Direction);
                else
                    lineUn = Line.CreateUnbound(duoBranchPipe.SecondPipe.CurveSourcePipe.Origin, duoBranchPipe.SecondPipe.CurveSourcePipe.Direction);

                setComRet = lineUn.Intersect(lineZ, out intArr);
                if (setComRet != SetComparisonResult.Overlap)
                    return null;
                var intBranchPipe = intArr.get_Item(0).XYZPoint;
                intMainPipe = mainPipe.CurveSourcePipe_Unbound.Project(intBranchPipe).XYZPoint;

                using (SubTransaction reSubTrans = new SubTransaction(Global.UIDoc.Document))
                {
                    reSubTrans.Start();
                    var verticalPipe = Common.Clone(mainPipe.SourcePipe) as Pipe;
                    verticalPipe.LookupParameter("Diameter").Set(diameter);
                    (verticalPipe.Location as LocationCurve).Curve = Line.CreateBound(intMainPipe, intBranchPipe);
                    retVal = verticalPipe;

                    reSubTrans.Commit();
                }

                return retVal;
            }
            catch (Exception)
            { }
            return null;
        }

        private static bool CreateTap(MEPCurve mepCurveSplit1, MEPCurve mepCurveSplit2)
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
                    var con = Common.GetConnectorClosestTo(mepCurveSplit2, p10);
                    var elbow = Global.UIDoc.Document.Create.NewTakeoffFitting(con, mepCurveSplit1);
                    Global.UIDoc.Document.Regenerate();
                }
                else
                {
                    var con = Common.GetConnectorClosestTo(mepCurveSplit2, p11);
                    var elbow = Global.UIDoc.Document.Create.NewTakeoffFitting(con, mepCurveSplit1);
                    Global.UIDoc.Document.Regenerate();
                }

                return true;
            }
            catch (System.Exception ex)
            {
                return false;
            }
        }

        private static void CreateTopTee(DuoBranchPipe duoBranchPipe, TwoLevelSmartDialogData dialogData)
        {
            double diamterPipe = dialogData.PipeSize * Common.mmToFT;

            Pipe pipe_1 = null;
            Pipe pipe_2 = null;
            Pipe verticalPipe = null;

            if (duoBranchPipe.FirstPipe == null)
                return;
            pipe_1 = duoBranchPipe.FirstPipe.SourcePipe;

            if (duoBranchPipe.SecondPipe != null)
                pipe_2 = duoBranchPipe.SecondPipe.SourcePipe;

            if (duoBranchPipe.VerticalPipe == null)
                return;
            verticalPipe = duoBranchPipe.VerticalPipe.SourcePipe;

            XYZ p0_sub1 = null;
            XYZ p1_sub1 = null;
            XYZ inters_sub1 = null;
            bool twoPipes = duoBranchPipe.SecondPipe != null ? true : false;

            GetInforBasePointTopTee(verticalPipe, pipe_1, twoPipes, out p0_sub1, out p1_sub1, out inters_sub1);
            if (duoBranchPipe.SecondPipe == null)
            {
                var temp_pipe1 = pipe_1;
                try
                {
                    var temp_pipe2_id = PlumbingUtils.BreakCurve(Global.UIDoc.Document, temp_pipe1.Id, inters_sub1);
                    var temp_pipe2 = Global.UIDoc.Document.GetElement(temp_pipe2_id) as Pipe;
                    pipe_1 = temp_pipe1;
                    pipe_2 = temp_pipe2;
                }
                catch (Exception)
                {
                    Global.UIDoc.Document.Delete(verticalPipe.Id);
                    return;
                }
            }

            GetInforBasePointTopTee(verticalPipe, pipe_1, true, out p0_sub1, out p1_sub1, out inters_sub1);

            XYZ p0_sub2 = null;
            XYZ p1_sub2 = null;
            XYZ inters_sub2 = null;

            GetInforBasePointTopTee(verticalPipe, pipe_2, true, out p0_sub2, out p1_sub2, out inters_sub2);

            if (inters_sub1.DistanceTo(inters_sub2) > 0.001)
            {
                return;
            }

            //Process for sub pipes
            var c5 = Common.GetConnectorClosestTo(verticalPipe, inters_sub1);

            double dMoi = 15 * Common.mmToFT;

            double d10 = 10 * Common.mmToFT;
            double dtemp = 2;

            if (CompareDouble(diamterPipe, pipe_1.Diameter) && CompareDouble(diamterPipe, pipe_2.Diameter))
            {
                using (SubTransaction reSubTrans = new SubTransaction(Global.UIDoc.Document))
                {
                    reSubTrans.Start();
                    //Connect to sub
                    (pipe_1.Location as LocationCurve).Curve = Line.CreateBound(p0_sub1, p1_sub1);
                    (pipe_2.Location as LocationCurve).Curve = Line.CreateBound(p0_sub2, p1_sub2);

                    //Connect
                    var c3 = Common.GetConnectorClosestTo(pipe_1, inters_sub1);
                    var c4 = Common.GetConnectorClosestTo(pipe_2, inters_sub1);

                    if (CreateTee(c3, c4, c5) == null)
                    {
                        reSubTrans.RollBack();
                    }
                    else
                        reSubTrans.Commit();
                }
            }
            else
            {
                using (SubTransaction reSubTrans = new SubTransaction(Global.UIDoc.Document))
                {
                    reSubTrans.Start();
                    Connector c3 = null;
                    Pipe pipe_moi_1 = null;

                    XYZ v = null;
                    if (CompareDouble(diamterPipe, pipe_1.Diameter) == false)
                    {
                        //Tao mot ong mồi
                        pipe_moi_1 = Common.Clone(verticalPipe) as Pipe;

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
                        (pipe_1.Location as LocationCurve).Curve = line;
                        c3 = Common.GetConnectorClosestTo(pipe_1, inters_sub1);

                        v = line.Direction;
                    }

                    Connector c4 = null;
                    Pipe pipe_moi_2 = null;

                    if (CompareDouble(diamterPipe, pipe_1.Diameter) == false)
                    {
                        //Tao mot ong mồi
                        pipe_moi_2 = Common.Clone(verticalPipe) as Pipe;

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
                        (pipe_2.Location as LocationCurve).Curve = Line.CreateBound(p0_sub2, p1_sub2);
                        c4 = Common.GetConnectorClosestTo(pipe_2, inters_sub1);
                    }

                    //Connect
                    FamilyInstance fitting = CreateFittingForMEPUtils.ct(c3, c4, c5);
                    if (fitting == null)
                    {
                        reSubTrans.RollBack/*Commit*/();
                    }
                    else
                    {
                        Connector main1 = null;
                        Connector main2 = null;
                        Connector tee = null;
                        Common.GetInfo(fitting, v, out main1, out main2, out tee);

                        var mep1 = GetPipeFromConnector(main1.AllRefs);
                        var mep2 = GetPipeFromConnector(main2.AllRefs);

                        var main1_p = main1.Origin;
                        var v_m_1 = main1.CoordinateSystem.BasisZ;

                        var main2_p = main2.Origin;
                        var v_m_2 = main2.CoordinateSystem.BasisZ;

                        FamilyInstance topReducer_1 = null;
                        FamilyInstance topReducer_2 = null;
                        //1
                        if (pipe_moi_1 != null)
                        {
                            if (mep1 != null && pipe_moi_1.Id == mep1.Id)
                            {
                                var line = Line.CreateUnbound(main1_p, v_m_1 * 10);

                                var p = line.Evaluate(d10, false);
                                (pipe_moi_1.Location as LocationCurve).Curve = Line.CreateBound(main1_p, p);

                                var futher = inters_sub1.DistanceTo(p0_sub1) > inters_sub1.DistanceTo(p1_sub1) ? p0_sub1 : p1_sub1;

                                ll(pipe_1, futher, p);

                                topReducer_1 = CreateFittingForMEPUtils.ctt(pipe_1, pipe_moi_1);
                            }
                            else if (mep2 != null && pipe_moi_1.Id == mep2.Id)
                            {
                                var line = Line.CreateUnbound(main2_p, v_m_2 * 10);

                                var p = line.Evaluate(d10, false);
                                (pipe_moi_1.Location as LocationCurve).Curve = Line.CreateBound(main2_p, p);

                                var futher = inters_sub1.DistanceTo(p0_sub1) > inters_sub1.DistanceTo(p1_sub1) ? p0_sub1 : p1_sub1;

                                ll(pipe_1, futher, p);

                                topReducer_1 = CreateFittingForMEPUtils.ctt(pipe_1, pipe_moi_1);
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

                                ll(pipe_2, futher, p);

                                topReducer_2 = CreateFittingForMEPUtils.ctt(pipe_2, pipe_moi_2);
                            }
                            else if (mep2 != null && pipe_moi_2.Id == mep2.Id)
                            {
                                var line = Line.CreateUnbound(main2_p, v_m_2 * 10);

                                var p = line.Evaluate(d10, false);
                                (pipe_moi_2.Location as LocationCurve).Curve = Line.CreateBound(main2_p, p);

                                var futher = inters_sub2.DistanceTo(p0_sub2) > inters_sub2.DistanceTo(p1_sub2) ? p0_sub2 : p1_sub2;

                                ll(pipe_2, futher, p);

                                topReducer_2 = CreateFittingForMEPUtils.ctt(pipe_2, pipe_moi_2);
                            }
                        }

                        if (dialogData.FlagAddNipple == true && dialogData.NippleFamily != null)
                        {
                            // Nipple Family
                            HandlerConnectAccessoryNipple(pipe_moi_1, topReducer_1, fitting, dialogData.NippleFamily);
                            HandlerConnectAccessoryNipple(pipe_moi_2, topReducer_2, fitting, dialogData.NippleFamily);
                        }
                        reSubTrans.Commit();
                    }
                }
            }
        }

        private static bool CompareDouble(double d1, double d2)
        {
            if (Math.Abs(d1 - d2) < 0.001)
                return true;

            return false;
        }

        private static bool GetInforBasePointTopTee(Pipe verticalPipe, Pipe branchPipe, bool twoPipe, out XYZ p0_sub, out XYZ p1_sub, out XYZ inters_sub)
        {
            inters_sub = verticalPipe.GetCurve().GetEndPoint(1);
            var curve_sub = branchPipe.GetCurve();
            if (twoPipe)
            {
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

        public static FamilyInstance CreateTee(Connector c3, Connector c4, Connector c5)
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

        private static Pipe GetPipeFromConnector(ConnectorSet cs)
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

        private static void ll(Pipe pipe, XYZ pOn, XYZ pOther)
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

        private static bool HandlerConnectAccessoryNipple(Pipe primerPipe_1, FamilyInstance reducer_1, FamilyInstance tee, FamilySymbol familySymbol)
        {
            try
            {
                Tuple<Connector, Connector> cntsNipple_1 = PipeToPairConnector(primerPipe_1, reducer_1, tee);

                if (cntsNipple_1.Item1 == null || cntsNipple_1.Item2 == null)
                    return false;

                Line lineNipple = Line.CreateBound(cntsNipple_1.Item1.Origin, cntsNipple_1.Item2.Origin);

                XYZ center = MiddlePoint(cntsNipple_1.Item1.Origin, cntsNipple_1.Item2.Origin);

                FamilyInstance nipple = Global.UIDoc.Document.Create.NewFamilyInstance(center, familySymbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

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
            catch (Exception)
            { }
            return false;
        }

        private static Tuple<Connector, Connector> PipeToPairConnector(Pipe pipe, FamilyInstance reducer, FamilyInstance tee)
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

        private static XYZ MiddlePoint(XYZ xYZ_1, XYZ xYZ_2) => new XYZ((xYZ_1.X + xYZ_2.X) * 0.5, (xYZ_1.Y + xYZ_2.Y) * 0.5, (xYZ_1.Z + xYZ_2.Z) * 0.5);

        private static void RotateLine(Document doc, FamilyInstance wye, Line axisLine)
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

        private static XYZ GetUnBoundIntersection(Line Line1, Line Line2)
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

        private static Pipe ProcessMainPipe(List<Pipe> mainPipes, Pipe processPipe, out bool flagSplit, out XYZ processIntPnt)
        {
            flagSplit = false;
            processIntPnt = null;
            try
            {
                Line flattenCurve_processPipe = Line.CreateUnbound(processPipe.GetFlattenCurve().GetEndPoint(0), processPipe.GetFlattenCurve().Direction);
                XYZ originPnt = processPipe.GetFlattenCurve().GetEndPoint(0);

                List<Tuple<Pipe, XYZ>> validMainPipes_Expand = new List<Tuple<Pipe, XYZ>>();
                List<Tuple<Pipe, XYZ>> validMainPipes_Real = new List<Tuple<Pipe, XYZ>>();

                Dictionary<Pipe, double> dictMainPipes_Expand = new Dictionary<Pipe, double>();
                Dictionary<Pipe, double> dictMainPipes_Real = new Dictionary<Pipe, double>();
                foreach (var mainPipe in mainPipes)
                {
                    XYZ intPnt;
                    Line flattenCurve_mainPipe = mainPipe.GetFlattenCurve();
                    if (RealityIntersect(flattenCurve_processPipe, flattenCurve_mainPipe, out intPnt))
                    {
                        validMainPipes_Real.Add(new Tuple<Pipe, XYZ>(mainPipe, intPnt));
                        dictMainPipes_Real.Add(mainPipe, originPnt.DistanceTo(intPnt));
                    }

                    Line flattenCurve_mainPipe_Expand = mainPipe.GetExpandFlattenCurve(700 * Common.mmToFT);
                    if (RealityIntersect(flattenCurve_processPipe, flattenCurve_mainPipe_Expand, out intPnt))
                    {
                        validMainPipes_Expand.Add(new Tuple<Pipe, XYZ>(mainPipe, intPnt));
                        dictMainPipes_Expand.Add(mainPipe, originPnt.DistanceTo(intPnt));
                    }
                }

                if (validMainPipes_Real.Count > 0)
                {
                    double minDist = dictMainPipes_Real.Min(item => item.Value);

                    var validDic = dictMainPipes_Real.FirstOrDefault(item => item.Value == minDist);

                    Pipe pipeNear = null;

                    foreach (var validDataMainPipe in validMainPipes_Real)
                    {
                        if (validDataMainPipe.Item1.Id != validDic.Key.Id)
                            continue;

                        var curve = (validDataMainPipe.Item1.Location as LocationCurve).Curve;

                        if (curve is Line == false)
                            continue;

                        var d = (curve as Line).Direction;

                        if (Common.IsParallel(d, XYZ.BasisZ, 0))
                            continue;

                        var project = curve.Project(validDataMainPipe.Item2);
                        if (project == null)
                            continue;

                        var p = project.XYZPoint;
                        processIntPnt = p;

                        if (p.DistanceTo(curve.GetEndPoint(0)) != 0 && p.DistanceTo(curve.GetEndPoint(1)) != 0)
                        {
                            flagSplit = true;

                            return validDataMainPipe.Item1;
                        }
                        else
                        {
                            pipeNear = validDataMainPipe.Item1;
                        }
                    }
                    return pipeNear;
                }
                else
                {
                    if (validMainPipes_Expand.Count <= 0)
                        return null;

                    double minDist = dictMainPipes_Expand.Min(item => item.Value);

                    var validDic = dictMainPipes_Expand.FirstOrDefault(item => item.Value == minDist);

                    Pipe pipeNear = null;

                    foreach (var validDataMainPipe in validMainPipes_Expand)
                    {
                        if (validDataMainPipe.Item1.Id != validDic.Key.Id)
                            continue;

                        var curve = (validDataMainPipe.Item1.Location as LocationCurve).Curve;

                        if (curve is Line == false)
                            continue;

                        var d = (curve as Line).Direction;

                        if (Common.IsParallel(d, XYZ.BasisZ, 0))
                            continue;

                        var project = curve.Project(validDataMainPipe.Item2);
                        if (project == null)
                            continue;

                        var p = project.XYZPoint;
                        processIntPnt = p;

                        if (p.DistanceTo(curve.GetEndPoint(0)) != 0 && p.DistanceTo(curve.GetEndPoint(1)) != 0)
                        {
                            flagSplit = true;

                            return validDataMainPipe.Item1;
                        }
                        else
                        {
                            pipeNear = validDataMainPipe.Item1;
                        }
                    }
                    return pipeNear;
                }
            }
            catch (Exception)
            { }
            return null;
        }

        public static void ProcessStartSidePipe(Pipe pipe, out Pipe pipe2, XYZ pOn, bool flagSplit, out bool isDauOngChinh)
        {
            isDauOngChinh = false;
            var curve = (pipe.Location as LocationCurve).Curve;

            //Create plane
            var p0 = curve.GetEndPoint(0);
            var p1 = curve.GetEndPoint(1);

            //Check co phai dau cuu hoa o gan dau cua ong ko : check trong pham vi 1m - 400mm
            double kc_mm = 400 /*1000*/;

            if ((double)GetParameterValueByName(pipe, "Diameter") / Common.mmToFT >= 90)
                kc_mm = 1100;
            else if ((double)GetParameterValueByName(pipe, "Diameter") / Common.mmToFT >= 50)
                kc_mm = 600;

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
                if (IsIntersect(p0) == false && !CheckPipeIsEnd(pipe, pOn))
                {
                    isDauOng = true;
                    far = 1;
                }
            }
            else if (d2 < km_ft)
            {
                if (IsIntersect(p1) == false && !CheckPipeIsEnd(pipe, pOn))
                {
                    isDauOng = true;
                    far = 0;
                }
            }

            Pipe pipe1 = pipe;
            pipe2 = null;

            if (flagSplit == true)
            {
                if (isDauOng == false)
                {
                    SplitPipe(pipe, pOn, out pipe1, out pipe2);
                }
                else if (far != -1)
                {
                    if (far == 1)
                        (pipe1.Location as LocationCurve).Curve = Line.CreateBound(pOn, p1);
                    else
                        (pipe1.Location as LocationCurve).Curve = Line.CreateBound(p0, pOn);
                }
            }
            else
            {
                if (isDauOng == false)
                { }
                else if (far != -1)
                {
                    isDauOngChinh = true;
                    if (far == 1)
                        (pipe1.Location as LocationCurve).Curve = Line.CreateBound(pOn, p1);
                    else
                        (pipe1.Location as LocationCurve).Curve = Line.CreateBound(p0, pOn);
                }
            }
        }

        public static bool IsIntersect(XYZ point)
        {
            double ft = 0.001;
            var solid = Common.CreateCylindricalVolume(point, ft, ft, false);
            if (solid != null)
            {
                var fittingBuilt = new ElementId(BuiltInCategory.OST_PipeFitting);
                FilteredElementCollector collector = new FilteredElementCollector(Global.UIDoc.Document);
                collector.OfClass(typeof(FamilyInstance));
                collector.OfCategoryId(fittingBuilt);
                collector.WherePasses(new ElementIntersectsSolidFilter(solid));

                if (collector.GetElementCount() != 0)
                {
                    return true;
                }
            }
            return false;
        }

        public static void SplitPipe(Pipe pipeOrigin, XYZ splitPoint, out Pipe pipe1, out Pipe pipe2)
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

        public static bool CheckPipeIsEnd(Pipe pipe, XYZ point)
        {
            var con = Common.GetConnectorClosestTo(pipe, point);

            return con.IsConnected;
        }

        private static dynamic GetParameterValueByName(Element elem, string paramName)
        {
            if (elem != null)
            {
                Parameter parameter = elem.LookupParameter(paramName);
                return GetParameterValue(parameter);
            }
            return null;
        }

        private static dynamic GetParameterValue(Parameter parameter)
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

        private static bool DivideCase_2(Pipe main, Pipe pipeSub, bool twoPipes, out XYZ inters_main, out XYZ p0_sub, out XYZ p1_sub, out XYZ inters_sub)
        {
            inters_main = null;
            p0_sub = null;
            p1_sub = null;
            inters_sub = null;

            var curve_main = main.GetCurve();
            var curve_main_2d = Line.CreateBound(Common.ToPoint2D(curve_main.GetEndPoint(0)), Common.ToPoint2D(curve_main.GetEndPoint(1)));
            var curve_main_expand = Line.CreateUnbound(curve_main.GetEndPoint(0), (curve_main as Line).Direction * 100);

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
            inters_main = curve_main_expand.Project(inters_sub).XYZPoint;
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
                        XYZ newStartPoint = CurveSourcePipe.GetEndPoint(0) + (SourcePipe_Cont_1.Origin - SourcePipe_Cont_2.Origin).Normalize() * 1000 * Common.mmToFT;
                        XYZ newEndPoint = CurveSourcePipe.GetEndPoint(1);

                        //newStartPoint = new XYZ(Math.Round(newStartPoint.X, 7), Math.Round(newStartPoint.Y, 7), Math.Round(newStartPoint.Z, 7));
                        //newEndPoint = new XYZ(Math.Round(newEndPoint.X, 7), Math.Round(newEndPoint.Y, 7), Math.Round(newEndPoint.Z, 7));
                        return Line.CreateBound(newStartPoint, newEndPoint);
                    }
                    else if (SourcePipe_Cont_1.IsConnected && !SourcePipe_Cont_2.IsConnected)
                    {
                        XYZ newStartPoint = CurveSourcePipe.GetEndPoint(0);
                        XYZ newEndPoint = CurveSourcePipe.GetEndPoint(1) + (SourcePipe_Cont_2.Origin - SourcePipe_Cont_1.Origin).Normalize() * 1000 * Common.mmToFT;

                        //newEndPoint = new XYZ(Math.Round(newEndPoint.X, 7), Math.Round(newEndPoint.Y, 7), Math.Round(newEndPoint.Z, 7));
                        //newStartPoint = new XYZ(Math.Round(newStartPoint.X, 7), Math.Round(newStartPoint.Y, 7), Math.Round(newStartPoint.Z, 7));
                        return Line.CreateBound(newStartPoint, newEndPoint);
                    }
                    else if (SourcePipe_Cont_1.IsConnected && SourcePipe_Cont_2.IsConnected)
                    {
                        return CurveSourcePipe;
                    }
                    else
                    {
                        XYZ newStartPoint = CurveSourcePipe.Evaluate(CurveSourcePipe.GetEndParameter(0) - 1000 * Common.mmToFT, false);
                        XYZ newEndPoint = CurveSourcePipe.Evaluate(CurveSourcePipe.GetEndParameter(1) + 1000 * Common.mmToFT, false);
                        return Line.CreateBound(newStartPoint, newEndPoint);
                    }
                }
                return null;
            }
        }

        public Line CurveSourcePipe_Extend_Round
        {
            get
            {
                if (SourcePipe != null && SourcePipe.IsValidObject && SourcePipe_Cont_1 != null && SourcePipe_Cont_2 != null)
                {
                    if (!SourcePipe_Cont_1.IsConnected && SourcePipe_Cont_2.IsConnected)
                    {
                        XYZ newStartPoint = CurveSourcePipe.GetEndPoint(0) + (SourcePipe_Cont_1.Origin - SourcePipe_Cont_2.Origin).Normalize() * 1000 * Common.mmToFT;
                        XYZ newEndPoint = CurveSourcePipe.GetEndPoint(1);

                        newStartPoint = new XYZ(Math.Round(newStartPoint.X, 2), Math.Round(newStartPoint.Y, 2), TruncateDouble(newStartPoint.Z));
                        newEndPoint = new XYZ(Math.Round(newEndPoint.X, 2), Math.Round(newEndPoint.Y, 2), TruncateDouble(newEndPoint.Z));
                        return Line.CreateBound(newStartPoint, newEndPoint);
                    }
                    else if (SourcePipe_Cont_1.IsConnected && !SourcePipe_Cont_2.IsConnected)
                    {
                        XYZ newStartPoint = CurveSourcePipe.GetEndPoint(0);
                        XYZ newEndPoint = CurveSourcePipe.GetEndPoint(1) + (SourcePipe_Cont_2.Origin - SourcePipe_Cont_1.Origin).Normalize() * 1000 * Common.mmToFT;

                        newEndPoint = new XYZ(Math.Round(newEndPoint.X, 2), Math.Round(newEndPoint.Y, 2), TruncateDouble(newEndPoint.Z));
                        newStartPoint = new XYZ(Math.Round(newStartPoint.X, 2), Math.Round(newStartPoint.Y, 2), TruncateDouble(newStartPoint.Z));
                        return Line.CreateBound(newStartPoint, newEndPoint);
                    }
                    else if (SourcePipe_Cont_1.IsConnected && SourcePipe_Cont_2.IsConnected)
                    {
                        return CurveSourcePipe;
                    }
                    else
                    {
                        XYZ newStartPoint = CurveSourcePipe.Evaluate(CurveSourcePipe.GetEndParameter(0) - 1000 * Common.mmToFT, false);
                        XYZ newEndPoint = CurveSourcePipe.Evaluate(CurveSourcePipe.GetEndParameter(1) + 1000 * Common.mmToFT, false);
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

        public Tuple<XYZ, XYZ> ExtendValidPnts { get; set; }

        public Line CuvreSourcePipe_FlattenExtend
        {
            get
            {
                if (SourcePipe != null && SourcePipe.IsValidObject && SourcePipe_Cont_1 != null && SourcePipe_Cont_2 != null)
                {
                    if (!SourcePipe_Cont_1.IsConnected && SourcePipe_Cont_2.IsConnected)
                    {
                        XYZ newStartPoint = CurveSourcePipe_Flatten.Evaluate(CurveSourcePipe_Flatten.GetEndParameter(0) - 600 * Common.mmToFT, false);
                        XYZ newEndPoint = CurveSourcePipe_Flatten.GetEndPoint(1);
                        ExtendValidPnts = new Tuple<XYZ, XYZ>(newStartPoint, null);
                        return Line.CreateBound(newStartPoint, newEndPoint);
                    }
                    else if (SourcePipe_Cont_1.IsConnected && !SourcePipe_Cont_2.IsConnected)
                    {
                        XYZ newStartPoint = CurveSourcePipe_Flatten.GetEndPoint(0);
                        XYZ newEndPoint = CurveSourcePipe_Flatten.Evaluate(CurveSourcePipe_Flatten.GetEndParameter(1) + 600 * Common.mmToFT, false);
                        ExtendValidPnts = new Tuple<XYZ, XYZ>(newEndPoint, null);
                        return Line.CreateBound(newStartPoint, newEndPoint);
                    }
                    else if (SourcePipe_Cont_1.IsConnected && SourcePipe_Cont_2.IsConnected)
                    {
                        ExtendValidPnts = new Tuple<XYZ, XYZ>(null, null);
                        return CurveSourcePipe_Flatten;
                    }
                    else
                    {
                        XYZ newStartPoint = CurveSourcePipe_Flatten.Evaluate(CurveSourcePipe_Flatten.GetEndParameter(0) - 600 * Common.mmToFT, false);
                        XYZ newEndPoint = CurveSourcePipe_Flatten.Evaluate(CurveSourcePipe_Flatten.GetEndParameter(1) + 600 * Common.mmToFT, false);
                        ExtendValidPnts = new Tuple<XYZ, XYZ>(newStartPoint, newEndPoint);
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

        public double TruncateDouble(double number)
        {
            return Math.Truncate(number * 10000) / 10000;
        }
    }

    public class DuoBranchPipe : ICloneable
    {
        public InforPie FirstPipe { get; set; }
        public InforPie SecondPipe { get; set; }
        public XYZ FlattenIntersectPoint { get; set; }
        public InforPie VerticalPipe { get; set; }

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
                    XYZ f_p1 = FirstPipe.SourcePipe_Cont_1.Origin.FlattenPoint();
                    XYZ f_p2 = FirstPipe.SourcePipe_Cont_2.Origin.FlattenPoint();
                    XYZ s_p1 = SecondPipe.SourcePipe_Cont_1.Origin.FlattenPoint();
                    XYZ s_p2 = SecondPipe.SourcePipe_Cont_2.Origin.FlattenPoint();
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
                    return FirstPipe.CurveSourcePipe_Flatten;
                }
                return null;
            }
        }

        public object Clone()
        {
            return this.MemberwiseClone();
        }
    }

    public class SourceMainPipe
    {
        public InforPie MainPipe { get; set; }
        public List<DuoBranchPipe> Branches { get; set; }
        public List<DuoBranchPipe> Branches_Special { get; set; }
        public XYZ Flatten_ExtendValidPnt_1 { get; set; }
        public XYZ Flatten_ExtendValidPnt_2 { get; set; }
        public DuoBranchPipe IntDuoBranch_1 { get; set; }
        public DuoBranchPipe IntDuoBranch_2 { get; set; }

        public SourceMainPipe(InforPie mainPipe)
        {
            MainPipe = mainPipe;
            Branches = new List<DuoBranchPipe>();
            Branches_Special = new List<DuoBranchPipe>();
            if (mainPipe.CuvreSourcePipe_FlattenExtend != null && mainPipe.ExtendValidPnts != null)
            {
                Flatten_ExtendValidPnt_1 = mainPipe.ExtendValidPnts.Item1;
                Flatten_ExtendValidPnt_2 = mainPipe.ExtendValidPnts.Item2;
            }
        }

        public void Initialize()
        {
            if (Branches_Special != null && Branches_Special.Count > 0)
            {
                if (Flatten_ExtendValidPnt_1 != null)
                {
                    try
                    {
                        XYZ closetPnt1 = FindClosestPoint(Flatten_ExtendValidPnt_1, Branches_Special);
                        IntDuoBranch_1 = Branches_Special.Where(item => Common.IsEqual(item.FlattenIntersectPoint, closetPnt1)).First();
                        if (IntDuoBranch_1 != null)
                        {
                            Branches_Special.Remove(IntDuoBranch_1);
                        }
                    }
                    catch (Exception)
                    { }
                }

                if (Flatten_ExtendValidPnt_2 != null)
                {
                    try
                    {
                        XYZ closetPnt2 = FindClosestPoint(Flatten_ExtendValidPnt_2, Branches_Special);
                        IntDuoBranch_2 = Branches_Special.Where(item => Common.IsEqual(item.FlattenIntersectPoint, closetPnt2)).First();
                        if (IntDuoBranch_2 != null)
                        {
                            Branches_Special.Remove(IntDuoBranch_2);
                        }
                    }
                    catch (Exception)
                    { }
                }
            }
        }

        public XYZ FindClosestPoint(XYZ origin, List<DuoBranchPipe> points)
        {
            if (points == null || points.Count <= 0)
                return null;

            double minDistance = double.MaxValue;
            XYZ closestPoint = null;

            foreach (var point in points)
            {
                double distance = origin.DistanceTo(point.FlattenIntersectPoint);
                if (distance < minDistance)
                {
                    minDistance = distance;
                    closestPoint = point.FlattenIntersectPoint;
                }
            }

            return closestPoint;
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