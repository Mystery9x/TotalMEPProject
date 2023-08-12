using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;
using System;
using TotalMEPProject.Ultis.StorageUtility;
using TotalMEPProject.Ultis;
using TotalMEPProject.SelectionFilters;

namespace TotalMEPProject.Commands.TotalMEP
{
    [Transaction(TransactionMode.Manual)]
    public class CmdPickLine : IExternalCommand
    {
        public static List<double> Offsets = new List<double>();

        public static List<Pipe> AllPipes = new List<Pipe>();

        public static bool IsRunning = false;

        public static bool isRedo_Undo = false;

        //public static Dictionary<MEPCurve, Curve> MEPCurves = new Dictionary<MEPCurve, Curve>();

        //public static List<MEPObjectData> _Datas = new List<MEPObjectData>();

        public static double MAX_REDUCER = 1000 * Common.mmToFT;
        public static double MAX_FSILON = 200 * Common.mmToFT;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
#if RELEASE
            if (App.CROK() != true)
                return Result.Failed;

#endif
            Global.UIApp = commandData.Application;
            Global.RVTApp = commandData.Application.Application;
            Global.UIDoc = commandData.Application.ActiveUIDocument;
            Global.AppCreation = commandData.Application.Application.Create;

            if (App.ShowCreateMEPForm() == false)
            {
                return Result.Cancelled;
            }

            return Result.Succeeded;
        }

        private static Transaction m_Transaction = null;
        private static SubTransaction m_Subtransaction = null;
        public static string _TransactionName = Guid.NewGuid().ToString();

        private static int Index_data = 0;

        //

        //public static bool ir()
        //{
        //    if (App.m_CreateMEPForm.IsCreateFitting_Elbow == true || App.m_CreateMEPForm.IsCreateFitting_Tee == true || App.m_CreateMEPForm.IsCreateFitting_Transition == true)
        //        return true;

        //    return false;
        //}

        //public static Result o()
        //{
        //    if (App.m_CreateMEPForm == null)
        //        return Result.Cancelled;

        //    //MEPCurves = new Dictionary<MEPCurve, Curve>();

        //    //_Datas.Clear();
        //    Index_data = 0;

        //    while (true)
        //    {
        //        // ValidationList();

        //        //                 if (IsRunning == false)
        //        //                 {
        //        //                     App.m_CreateMEPForm._WindowHookcs.HookMainApp();
        //        //                 }

        //        IsRunning = true;

        //        m_Transaction = new Transaction(Global.UIDoc.Document, _TransactionName);
        //        m_Transaction.Start();

        //        m_Subtransaction = new SubTransaction(Global.UIDoc.Document);
        //        m_Subtransaction.Start();

        //        //Events
        //        //Global.m_rvtApp.DocumentChanged += new EventHandler<DocumentChangedEventArgs>(OnDocumentChanged);

        //        Reference pickedObj = null;
        //        var pickLineFilter = new SelectionFilters.PickLineFiler();

        //        Curve curve = null;

        //        bool error = false;
        //        try
        //        {
        //            pickedObj = Global.UIDoc.Selection.PickObject(ObjectType.PointOnElement, pickLineFilter, "Pick multiple lines: ");

        //            curve = Line.CreateBound(pickLineFilter.Start, pickLineFilter.End);

        //            Index_data++;
        //        }
        //        catch (System.Exception ex)
        //        {
        //            error = true;

        //            m_Subtransaction.RollBack();
        //            m_Transaction.RollBack();

        //            //RemoveHiden();

        //            if (App.m_CreateMEPForm != null)
        //                App.m_CreateMEPForm.SetCancelText("Close");

        //            IsRunning = false;
        //            //Global.m_rvtApp.DocumentChanged -= new EventHandler<DocumentChangedEventArgs>(OnDocumentChanged);

        //            //                     if (App.m_CreateMEPForm._WindowHookcs != null)
        //            //                     {
        //            //                         App.m_CreateMEPForm._WindowHookcs.UnhookMainApp();
        //            //                     }

        //            return Result.Succeeded;
        //        }
        //        finally
        //        {
        //            if (error == false)
        //            {
        //                if (m_Subtransaction.HasStarted())
        //                {
        //                    pickLineFilter.RemoveTemp();

        //                    pickLineFilter.RemoveDetailCurveItems();

        //                    m_Subtransaction.Commit();
        //                }

        //                if (m_Transaction.HasStarted())
        //                {
        //                    m_Transaction.Commit();
        //                }
        //            }
        //        }

        //        if (pickedObj == null || curve == null)
        //        {
        //            continue;
        //        }

        //        //Parameters
        //        var elementTypeId = App.m_CreateMEPForm.FamilyType;

        //        var find = Offsets.Find(item => item == App.m_CreateMEPForm.Offset);
        //        if (find == 0.0)
        //            Offsets.Add(App.m_CreateMEPForm.Offset);

        //        ElementId systemTypeId = ElementId.InvalidElementId;

        //        var levelId = App.m_CreateMEPForm.Level;

        //        double dOffset = 0;

        //        BuiltInCategory builtInCategory = BuiltInCategory.INVALID;

        //        if (App.m_CreateMEPForm.MEPType_ == SelectionFilters.MEPType.Pipe)
        //        {
        //            dOffset = (App.m_CreateMEPForm.MEPSize_ as MEPSize).OuterDiameter;

        //            systemTypeId = App.m_CreateMEPForm.SystemType;

        //            builtInCategory = BuiltInCategory.OST_PipeCurves;
        //        }
        //        else if (App.m_CreateMEPForm.MEPType_ == SelectionFilters.MEPType.Conduit)
        //        {
        //            dOffset = (App.m_CreateMEPForm.MEPSize_ as ConduitSize).OuterDiameter;

        //            builtInCategory = BuiltInCategory.OST_Conduit;
        //        }
        //        else if (App.m_CreateMEPForm.MEPType_ == SelectionFilters.MEPType.Round_Duct)
        //        {
        //            dOffset = (App.m_CreateMEPForm.MEPSize_ as MEPSize).NominalDiameter/*OuterDiameter*/;
        //            systemTypeId = App.m_CreateMEPForm.SystemType;

        //            builtInCategory = BuiltInCategory.OST_DuctCurves;
        //        }
        //        else if (App.m_CreateMEPForm.MEPType_ == SelectionFilters.MEPType.Oval_Duct || App.m_CreateMEPForm.MEPType_ == SelectionFilters.MEPType.Rectangular_Duct)
        //        {
        //            dOffset = App.m_CreateMEPForm.MEP_Width * Common.mmToFT;

        //            systemTypeId = App.m_CreateMEPForm.SystemType;

        //            builtInCategory = BuiltInCategory.OST_DuctCurves;
        //        }
        //        else
        //        {
        //            dOffset = App.m_CreateMEPForm.MEP_Width * Common.mmToFT;

        //            builtInCategory = BuiltInCategory.OST_CableTray;
        //        }

        //        var HorizontalEnum_ = App.m_CreateMEPForm.HorizontalEnum_;
        //        int hor = 0;

        //        var VerticalEnum_ = App.m_CreateMEPForm.VerticalEnum_;

        //        m_Transaction.Start();

        //        try
        //        {
        //            Reference mLine = pickedObj;
        //            var start = curve.GetEndPoint(0);
        //            var end = curve.GetEndPoint(1);

        //            if (HorizontalEnum_ != HorizontalEnum.Center)
        //            {
        //                var data = pickLineFilter.PickElements.Where(item => item.Key == pickedObj.ElementId).FirstOrDefault();
        //                var gp = pickedObj.GlobalPoint;

        //                var windowPoint = data.Value;

        //                //Check side
        //                //var side = Common.GetSide(start, end, windowPoint);

        //                var offsets = Common.CreateOffsetCurve(curve, dOffset / 2, true);
        //                //if (offsets.Count != 2)
        //                //    continue; //check

        //                //Get line with fatherest point
        //                Curve curve1 = offsets[0];
        //                Curve curve2 = offsets[1];

        //                Curve curve_left = curve1;
        //                Curve curve_right = curve2;

        //                var d1 = curve1.Distance(windowPoint);
        //                var d2 = curve2.Distance(windowPoint);

        //                if (d1 >/*<*/ d2)
        //                {
        //                    hor = 1; //Left
        //                }
        //                else
        //                {
        //                    hor = 2; //Right
        //                }

        //                if (d1 </*>*/ d2)
        //                {
        //                    start = curve_left.GetEndPoint(0);
        //                    end = curve_left.GetEndPoint(1);
        //                }
        //                else
        //                {
        //                    start = curve_right.GetEndPoint(0);
        //                    end = curve_right.GetEndPoint(1);
        //                }
        //            }

        //            if (App.m_CreateMEPForm.Slope_ != Slope.Off)
        //            {
        //                var slopeValue = App.m_CreateMEPForm.SlopeValue_;

        //                var z = (start - end).GetLength() * slopeValue / 100;

        //                end = new XYZ(end.X, end.Y, end.Z + (App.m_CreateMEPForm.Slope_ == Slope.Down ? z * -1 : z));
        //            }

        //            if (start != null && end != null)
        //            {
        //                //var mepNew = CreateFittingForMEPCommand.CreateMEP(mepTemp,start, end);

        //                var mepNew = err(start, end, elementTypeId, systemTypeId, levelId);

        //                if (mepNew == null)
        //                    continue;

        //                //Save entity
        //                string backup = string.Format("{0}|{1}|{2}", pickLineFilter.Curve_selected.GetType().ToString(),
        //                    start /*pickLineFilter.Curve_selected.GetEndPoint(0)*/,
        //                    end/* pickLineFilter.Curve_selected.GetEndPoint(1)*/);

        //                //Add to storage
        //                bool result = StorageUtility.AddEntity(mepNew, StorageUtility.m_CreatedFrom_Guild, StorageUtility.m_CreatedFrom, backup);

        //                //Comments
        //                var commentPara = mepNew.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
        //                if (commentPara != null)
        //                {
        //                    commentPara.Set(App.m_CreateMEPForm.Comments_);
        //                }

        //                //Set parameter
        //                double height = 0; //inch
        //                double width = 0;
        //                if (App.m_CreateMEPForm.MEPType_ == SelectionFilters.MEPType.Pipe || App.m_CreateMEPForm.MEPType_ == SelectionFilters.MEPType.Round_Duct)
        //                {
        //                    width = (App.m_CreateMEPForm.MEPSize_ as MEPSize).NominalDiameter;
        //                    mepNew.LookupParameter("Diameter").Set(width);

        //                    height = (App.m_CreateMEPForm.MEPSize_ as MEPSize).NominalDiameter;
        //                }
        //                else if (App.m_CreateMEPForm.MEPType_ == SelectionFilters.MEPType.Conduit)
        //                {
        //                    width = (App.m_CreateMEPForm.MEPSize_ as ConduitSize).NominalDiameter;
        //                    mepNew.LookupParameter("Diameter(Trade Size)").Set(width);

        //                    height = (App.m_CreateMEPForm.MEPSize_ as ConduitSize).NominalDiameter;
        //                }
        //                else
        //                {
        //                    width = App.m_CreateMEPForm.MEP_Width * Common.mmToFT;
        //                    mepNew.LookupParameter("Width").Set(width);
        //                    mepNew.LookupParameter("Height").Set(App.m_CreateMEPForm.MEP_Height * Common.mmToFT);

        //                    height = App.m_CreateMEPForm.MEP_Height * Common.mmToFT;
        //                }

        //                if (App.m_CreateMEPForm.MEPType_ == SelectionFilters.MEPType.CableTray || App.m_CreateMEPForm.MEPType_ == SelectionFilters.MEPType.Conduit && App.m_CreateMEPForm.ServiceType != string.Empty)
        //                {
        //                    mepNew.LookupParameter("Service Type").Set(App.m_CreateMEPForm.ServiceType);
        //                }

        //                //Swap
        //                if (HorizontalEnum_ != HorizontalEnum.Center && App.m_CreateMEPForm.Swap == true)
        //                    mepNew.LookupParameter("Horizontal Justification").Set(0); //Center
        //                else
        //                    mepNew.LookupParameter("Horizontal Justification").Set(hor);

        //                int ver = VerticalEnum_ == VerticalEnum.Middle ? 0 : (VerticalEnum_ == VerticalEnum.Bottom ? 1 : 2);
        //                mepNew.LookupParameter("Vertical Justification").Set(ver);

        //                double offset = App.m_CreateMEPForm.Offset;
        //                double haft = (height / Common.mmToFT) / 2; //mm

        //                if (VerticalEnum_ == VerticalEnum.Bottom)
        //                {
        //                    offset += haft;
        //                }
        //                else if (VerticalEnum_ == VerticalEnum.Top)
        //                {
        //                    offset -= haft;
        //                }
        //                Common.SetOffset(mepNew, offset);

        //                //Create fitting

        //                if (App.m_CreateMEPForm.Slope_ == Slope.Off)
        //                {
        //                    if (ir() == true)
        //                    {
        //                        var lastMEP = gg(mepNew.Id, mepNew.GetCurve(), elementTypeId, systemTypeId, width, builtInCategory, true);

        //                        if (lastMEP != null && mepNew != null)
        //                        {
        //                            //Curve goc
        //                            var isDauMut_CurvePicked = ee(mepNew, lastMEP);

        //                            var lastC = (lastMEP.Location as LocationCurve).Curve;

        //                            var intersection = ff(lastMEP, mepNew);
        //                            var isParalel = Common.IsParallel(lastMEP, mepNew);

        //                            if (intersection)
        //                            {
        //                                bool isCreateElbow = false;

        //                                if (isParalel == false && isDauMut_CurvePicked == true)
        //                                {
        //                                    //var pt = CreateFittingForMEPUtils.IsContinue(lastMEP, mepNew);

        //                                    isCreateElbow = true;
        //                                }

        //                                if (isCreateElbow)
        //                                    CreateFittingForMEPUtils.cb(lastMEP, mepNew);
        //                                else
        //                                {
        //                                    bool isTap = Common.IsTap(lastMEP);
        //                                    if (isTap && isDauMut_CurvePicked == false && ger(lastMEP, mepNew))
        //                                    {
        //                                        CreateFittingForMEPUtils.se(lastMEP, mepNew);
        //                                    }
        //                                    else if (isParalel && App.m_CreateMEPForm.IsCreateFitting_Transition && isDauMut_CurvePicked)
        //                                    {
        //                                        CreateFittingForMEPUtils.ctt(lastMEP, mepNew, true);
        //                                    }
        //                                    else
        //                                    {
        //                                        CreateFittingForMEPUtils.fe(lastMEP, mepNew);
        //                                    }
        //                                }
        //                            }
        //                            else
        //                            {
        //                                if (Common.IsParallel(lastMEP, mepNew) == true)
        //                                {
        //                                    if (App.m_CreateMEPForm.IsCreateFitting_Transition)
        //                                    {
        //                                        var same = CreateFittingForMEPUtils.iss(lastMEP, mepNew);
        //                                        if (same == true)
        //                                        {
        //                                            //Noi lai thanh 1 pipe
        //                                            CreateFittingForMEPUtils.dd(lastMEP, mepNew, true);
        //                                        }
        //                                        else
        //                                            CreateFittingForMEPUtils.ctt(lastMEP, mepNew, true);
        //                                    }
        //                                }
        //                                else
        //                                {
        //                                    if (mepNew is Duct || mepNew is CableTray)
        //                                    {
        //                                        bool isCreateElbow = false;
        //                                        if (isParalel == false && isDauMut_CurvePicked == true)
        //                                            isCreateElbow = true;

        //                                        if (isCreateElbow)
        //                                        {
        //                                            if (App.m_CreateMEPForm.IsCreateFitting_Elbow)
        //                                            {
        //                                                CreateFittingForMEPUtils.cb(lastMEP, mepNew);
        //                                            }
        //                                        }
        //                                        else
        //                                        {
        //                                            bool isTap = Common.IsTap(lastMEP);
        //                                            if (isTap && ger(lastMEP, mepNew))
        //                                            {
        //                                                CreateFittingForMEPUtils.se(lastMEP, mepNew);
        //                                            }
        //                                            else if (isParalel && App.m_CreateMEPForm.IsCreateFitting_Transition && isDauMut_CurvePicked)
        //                                            {
        //                                                CreateFittingForMEPUtils.ctt(lastMEP, mepNew, true);
        //                                            }
        //                                            else if (App.m_CreateMEPForm.IsCreateFitting_Tee)
        //                                                CreateFittingForMEPUtils.fe(lastMEP, mepNew);
        //                                        }
        //                                    }
        //                                }
        //                                //Tam thoi ko tao elbow
        //                                //                                         else if (App.m_CreateMEPForm.IsCreateFitting_Elbow)
        //                                //                                         {
        //                                //                                             CreateFittingForMEPUtils.CreateElbow(lastMEP, mepNew);
        //                                //                                         }
        //                            }
        //                        }
        //                        else
        //                        {
        //                            var newCurve = (mepNew.Location as LocationCurve).Curve;
        //                            er(mepNew, builtInCategory, elementTypeId, systemTypeId);
        //                        }
        //                    }
        //                }
        //                else
        //                {
        //                    MEPCurve mep = null;
        //                    mep = gg(mepNew.Id, mepNew.GetCurve(), elementTypeId, systemTypeId, width, builtInCategory, true /*false*/);

        //                    if (mep != null && mep as Pipe != null && mepNew as Pipe != null)
        //                    {
        //                        gr(mep as Pipe, mepNew as Pipe, mepNew.GetCurve(), elementTypeId, systemTypeId, levelId);
        //                    }
        //                }
        //            }
        //            m_Transaction.Commit();

        //            //Global.m_rvtApp.DocumentChanged -= new EventHandler<DocumentChangedEventArgs>(OnDocumentChanged);
        //        }
        //        catch (System.Exception ex)
        //        {
        //            m_Transaction.RollBack();
        //        }
        //    }
        //}

        //private static bool ee(Element element1, Element element2)
        //{
        //    ////BCB: Get tu entity . REM : LỖI TRƯỜNG HỢP, NGƯỜI DÙNG TỰ TAY KÉO DÀI PIPE RA
        //    //             var curve1 = GetCurve(element1);
        //    //             var curve2 = GetCurve(element2);

        //    var curve1 = (element1.Location as LocationCurve).Curve;
        //    var curve2 = (element2.Location as LocationCurve).Curve;

        //    if (curve1 == null || curve2 == null)
        //        return false;

        //    var pt = CreateFittingForMEPUtils.icc/*IsContinue_DauMut*/(curve1, curve2);
        //    if (pt == null)
        //        return false;
        //    else
        //        return true;
        //}

        //private static Curve gs(Element element)
        //{
        //    XYZ p0 = null;
        //    XYZ p1 = null;
        //    var v = ge(element, out p0, out p1);

        //    if (v == false)
        //        return null;

        //    return Line.CreateBound(p0, p1);
        //}

        //private static bool ge(Element element, out XYZ start, out XYZ end)
        //{
        //    start = null;
        //    end = null;

        //    var schema = Autodesk.Revit.DB.ExtensibleStorage.Schema.Lookup(StorageUtility.m_CreatedFrom_Guild);

        //    //Get created by category
        //    if (schema == null)
        //        return false;

        //    object result = StorageUtility.GetValue(element, schema, StorageUtility.m_CreatedFrom, typeof(string));

        //    if (result == null)
        //        return false;

        //    var str = result as string;

        //    if (str == null)
        //        return false;

        //    var splits = str.Split('|').ToList();

        //    if (splits.Count != 3)
        //        return false;

        //    var s1 = splits[1];
        //    s1 = s1.Replace("(", "");
        //    s1 = s1.Replace(")", "");

        //    var splits1 = s1.Split(',').ToList();
        //    if (splits1.Count != 3)
        //        return false;

        //    start = new XYZ(Convert.ToDouble(splits1[0].Trim()), Convert.ToDouble(splits1[1].Trim()), Convert.ToDouble(splits1[2].Trim()));

        //    var s2 = splits[2];
        //    s2 = s2.Replace("(", "");
        //    s2 = s2.Replace(")", "");

        //    var splits2 = s2.Split(',').ToList();
        //    if (splits2.Count != 3)
        //        return false;

        //    end = new XYZ(Convert.ToDouble(splits2[0].Trim()), Convert.ToDouble(splits2[1].Trim()), Convert.ToDouble(splits2[2].Trim()));

        //    return true;
        //}

        //private static bool ggg(MEPCurve lastMEP, MEPCurve mep)
        //{
        //    var lastC = lastMEP.GetCurve();
        //    var curve = mep.GetCurve();

        //    var p0 = new XYZ(lastC.GetEndPoint(0).X, lastC.GetEndPoint(0).Y, curve.GetEndPoint(0).Z);
        //    var p1 = new XYZ(lastC.GetEndPoint(1).X, lastC.GetEndPoint(1).Y, curve.GetEndPoint(0).Z);
        //    var lastCurve = Line.CreateBound(p0, p1);

        //    double fsilon = 50;// 5 * Common.mmToFT;

        //    bool intersection = true;
        //    if (curve.Intersect(lastCurve) != SetComparisonResult.Overlap)
        //    {
        //        intersection = false;

        //        //Check gan sat

        //        //Check same offset
        //        var offset1 = Utils.GetOffset(lastMEP);// lastMEP.LookupParameter("Offset").AsDouble();
        //        var offset2 = Utils.GetOffset(mep);//mep.LookupParameter("Offset").AsDouble();

        //        if (Math.Abs(offset1 - offset2) < Common._sixteenth)
        //        {
        //            //Expand
        //            IntersectionResultArray resultArray = new IntersectionResultArray();
        //            var curveExpand = Line.CreateBound(curve.Evaluate(curve.GetEndParameter(0) - fsilon, false), curve.Evaluate(curve.GetEndParameter(1) + fsilon, false));

        //            curveExpand = Line.CreateBound(Common.ToPoint2D(curveExpand.GetEndPoint(0), curve.GetEndPoint(0).Z), Common.ToPoint2D(curveExpand.GetEndPoint(1), curve.GetEndPoint(0).Z));
        //            if (curveExpand.Intersect(lastCurve, out resultArray) == SetComparisonResult.Overlap)
        //            {
        //                intersection = true;
        //            }
        //        }
        //    }

        //    return intersection;
        //}

        //private static bool ger(MEPCurve lastMEP, MEPCurve mep)
        //{
        //    var lastC = lastMEP.GetCurve();
        //    var curve = mep.GetCurve();

        //    //Êxtend
        //    double fsilon1 = 300 * Common.mmToFT;

        //    var p0 = curve.Evaluate(-fsilon1, false);
        //    var p1 = curve.Evaluate(curve.Length + fsilon1, false);
        //    curve = Line.CreateBound(p0, p1);

        //    p0 = new XYZ(lastC.GetEndPoint(0).X, lastC.GetEndPoint(0).Y, curve.GetEndPoint(0).Z);
        //    p1 = new XYZ(lastC.GetEndPoint(1).X, lastC.GetEndPoint(1).Y, curve.GetEndPoint(0).Z);
        //    var lastCurve = Line.CreateBound(p0, p1);

        //    double fsilon = 1 * Common.mmToFT;

        //    //Check intersection curve
        //    bool intersection = true;
        //    if (curve.Intersect(lastCurve) != SetComparisonResult.Overlap)
        //    {
        //        intersection = false;
        //    }

        //    return intersection;
        //}

        //private static bool ff(MEPCurve lastMEP, MEPCurve mep)
        //{
        //    var lastC = lastMEP.GetCurve();
        //    var curve = mep.GetCurve();

        //    var p0 = new XYZ(lastC.GetEndPoint(0).X, lastC.GetEndPoint(0).Y, curve.GetEndPoint(0).Z);
        //    var p1 = new XYZ(lastC.GetEndPoint(1).X, lastC.GetEndPoint(1).Y, curve.GetEndPoint(0).Z);
        //    var lastCurve = Line.CreateBound(p0, p1);

        //    double fsilon = 1 * Common.mmToFT;

        //    //Check intersection curve
        //    bool intersection = true;
        //    if (curve.Intersect(lastCurve) != SetComparisonResult.Overlap)
        //    {
        //        intersection = false;

        //        //Check gan sat

        //        //////////////////////////////////////////////////////////////////////////
        //        //Check same offset
        //        var offset1 = Utils.GetOffset(lastMEP);//lastMEP.LookupParameter("Offset").AsDouble();
        //        var offset2 = Utils.GetOffset(mep);// mep.LookupParameter("Offset").AsDouble();

        //        if (Math.Abs(offset1 - offset2) < Common._sixteenth)
        //        {
        //            //Expand
        //            IntersectionResultArray resultArray = new IntersectionResultArray();
        //            var curveExpand = Line.CreateBound(curve.Evaluate(curve.GetEndParameter(0) - fsilon, false), curve.Evaluate(curve.GetEndParameter(1) + fsilon, false));
        //            if (curveExpand.Intersect(lastCurve, out resultArray) == SetComparisonResult.Overlap)
        //            {
        //                intersection = true;

        //                var pInter = resultArray.get_Item(0).XYZPoint;

        //                //Change lenght
        //                var pN0 = curve.GetEndPoint(0);
        //                var pN1 = curve.GetEndPoint(1);

        //                var d0 = curve.GetEndPoint(0).DistanceTo(pInter);
        //                var d1 = curve.GetEndPoint(1).DistanceTo(pInter);

        //                if (d0 < d1)
        //                {
        //                    pN0 = pInter;
        //                }
        //                else
        //                {
        //                    pN1 = pInter;
        //                }

        //                (mep.Location as LocationCurve).Curve = Line.CreateBound(pN0, pN1);
        //            }
        //            else
        //            {
        //                //lastCurve = Line.CreateBound(lastCurve.Evaluate(lastCurve.GetEndParameter(0) - fsilon, false), lastCurve.Evaluate(lastCurve.GetEndParameter(1) + fsilon, false));
        //            }
        //        }
        //        //////////////////////////////////////////////////////////////////////////
        //    }

        //    //             var lastCurve2d = Line.CreateBound(Common.ToPoint2D(lastC.GetEndPoint(0)), Common.ToPoint2D(lastC.GetEndPoint(1)));
        //    //             var curve2d = Line.CreateBound(Common.ToPoint2D(curve.GetEndPoint(0)), Common.ToPoint2D(curve.GetEndPoint(1)));
        //    //
        //    //             //Check intersection curve
        //    //             bool intersection = true;
        //    //             if (curve2d.Intersect(lastCurve2d) != SetComparisonResult.Overlap)
        //    //                 intersection = false;

        //    return intersection;
        //}

        //public static void er(MEPCurve mepNew, BuiltInCategory builtInCategory, ElementId elementTypeId, ElementId systemTypeId)
        //{
        //    var newCurve = (mepNew.Location as LocationCurve).Curve;
        //    XYZ start = newCurve.GetEndPoint(0);

        //    XYZ end = newCurve.GetEndPoint(1);
        //    //Find nearest
        //    var v = (end - start).Normalize();

        //    //Check
        //    Type type = null;
        //    if (mepNew is Pipe)
        //    {
        //        type = typeof(Pipe);
        //    }
        //    else if (mepNew is Conduit)
        //    {
        //        type = typeof(Conduit);
        //    }
        //    else if (mepNew is Duct)
        //    {
        //        type = typeof(Duct);
        //        double h = 0;
        //        if (App.m_CreateMEPForm.MEPType_ == MEPType.Round_Duct)
        //        {
        //            h = mepNew.LookupParameter("Diameter").AsDouble();
        //        }
        //        else
        //            h = mepNew.LookupParameter("Height").AsDouble();

        //        var para = mepNew.get_Parameter(BuiltInParameter.RBS_CURVE_VERT_OFFSET_PARAM);
        //        if (para == null)
        //            return;

        //        var value = para.AsInteger();
        //        if (value == 1) //Bottom
        //        {
        //            start = new XYZ(start.X, start.Y, start.Z + h / 2);
        //        }
        //        else if (value == 2) //Top
        //        {
        //            start = new XYZ(start.X, start.Y, start.Z - h / 2);
        //        }
        //    }
        //    else if (mepNew is CableTray)
        //    {
        //        type = typeof(CableTray);
        //        double h = 0;

        //        h = mepNew.LookupParameter("Height").AsDouble();

        //        var para = mepNew.get_Parameter(BuiltInParameter.RBS_CURVE_VERT_OFFSET_PARAM);
        //        if (para == null)
        //            return;

        //        var value = para.AsInteger();
        //        if (value == 1) //Bottom
        //        {
        //            start = new XYZ(start.X, start.Y, start.Z + h / 2);
        //        }
        //        else if (value == 2) //Top
        //        {
        //            start = new XYZ(start.X, start.Y, start.Z - h / 2);
        //        }
        //    }

        //    var elements = Common.FindElementsByDirection(Global.m_uiDoc.Document, mepNew.Id, elementTypeId, systemTypeId, start, v, builtInCategory, true, type);
        //    if (elements.Count != 0)
        //    {
        //        var first = elements[0];

        //        if (first as MEPCurve != null)
        //        {
        //            var curveNear = ((first as MEPCurve).Location as LocationCurve).Curve;

        //            var v2 = curveNear.GetEndPoint(1) - curveNear.GetEndPoint(0);

        //            if (Utils.IsParallel(v, v2, Common.fsilon))
        //            {
        //                if (App.m_CreateMEPForm.IsCreateFitting_Transition)
        //                {
        //                    var same = CreateFittingForMEPUtils.iss(first as MEPCurve, mepNew);
        //                    if (same == false)
        //                        CreateFittingForMEPUtils.ctt(first as MEPCurve, mepNew, true);
        //                    else
        //                    {
        //                        //Noi lai thanh 1
        //                        CreateFittingForMEPUtils.dd(first as MEPCurve, mepNew, true);
        //                    }
        //                }
        //            }
        //        }
        //    }

        //    elements = Common.FindElementsByDirection(Global.m_uiDoc.Document, mepNew.Id, elementTypeId, systemTypeId, end, v.Negate(), builtInCategory, true, type);
        //    if (elements.Count != 0)
        //    {
        //        var first = elements[0];

        //        if (first as MEPCurve != null)
        //        {
        //            var curveNear = ((first as MEPCurve).Location as LocationCurve).Curve;

        //            var v2 = curveNear.GetEndPoint(1) - curveNear.GetEndPoint(0);

        //            if (Utils.IsParallel(v, v2, Common.fsilon))
        //            {
        //                if (App.m_CreateMEPForm.IsCreateFitting_Transition)
        //                {
        //                    var same = CreateFittingForMEPUtils.iss(first as MEPCurve, mepNew);
        //                    if (same == false)
        //                        CreateFittingForMEPUtils.ctt(first as MEPCurve, mepNew, true);
        //                    else
        //                    {
        //                        //Noi lai thanh 1
        //                        CreateFittingForMEPUtils.dd(first as MEPCurve, mepNew, true);
        //                    }
        //                }
        //            }
        //        }
        //    }
        //}

        //private static double gr(Curve curve1, Curve curve2)
        //{
        //    var p0 = curve1.GetEndPoint(0);
        //    var p1 = curve1.GetEndPoint(0);

        //    var p2 = curve2.GetEndPoint(0);
        //    var p3 = curve2.GetEndPoint(0);

        //    var dMin = double.MaxValue;
        //    if (p0.DistanceTo(p2) < dMin)
        //        dMin = p0.DistanceTo(p2);

        //    if (p0.DistanceTo(p3) < dMin)
        //        dMin = p0.DistanceTo(p3);

        //    if (p1.DistanceTo(p2) < dMin)
        //        dMin = p0.DistanceTo(p2);

        //    if (p1.DistanceTo(p3) < dMin)
        //        dMin = p0.DistanceTo(p2);

        //    return dMin;
        //}

        //private static bool ge(Curve curveMain, Curve pickCurve, out XYZ pClose, out int index, out bool isTee)
        //{
        //    double fsilon_1000 = 1000;
        //    double fsilon_10 = 10;

        //    pClose = null;
        //    isTee = false;
        //    index = -1;

        //    var pM0 = curveMain.GetEndPoint(0);
        //    var pM1 = curveMain.GetEndPoint(1);

        //    var startP = pickCurve.GetEndPoint(0);
        //    var endP = pickCurve.GetEndPoint(1);

        //    var main2d = Line.CreateBound(Common.ToPoint2D(pM0), Common.ToPoint2D(pM1));
        //    var current2d = Line.CreateBound(Common.ToPoint2D(startP), Common.ToPoint2D(endP));

        //    //expand
        //    var main_expand = Line.CreateBound(main2d.Evaluate(-fsilon_1000, false), main2d.Evaluate(main2d.GetEndParameter(1) + fsilon_1000, false));
        //    var current2d_expand = Line.CreateBound(current2d.Evaluate(-fsilon_1000, false), current2d.Evaluate(current2d.GetEndParameter(1) + fsilon_1000, false));

        //    IntersectionResultArray resultArray = new IntersectionResultArray();
        //    var interResult = current2d_expand.Intersect(main_expand, out resultArray);

        //    XYZ pClose2d = null;
        //    double fsilon = 0.1;
        //    if (interResult == SetComparisonResult.Overlap && resultArray.Size != 0)
        //    {
        //        pClose2d = resultArray.get_Item(0).XYZPoint;
        //    }
        //    else //if (interResult == SetComparisonResult.Equal)
        //    {
        //        //Song song va canh nhau
        //        var pM02d = Common.ToPoint2D(pM0);
        //        var pM12d = Common.ToPoint2D(pM1);

        //        var startP2d = Common.ToPoint2D(startP);
        //        var endP2d = Common.ToPoint2D(endP);

        //        var d = pM02d.DistanceTo(startP2d);
        //        if (d < fsilon)
        //        {
        //            pClose = pM0;
        //            index = 0;
        //            return true;
        //        }

        //        d = pM02d.DistanceTo(endP2d);
        //        if (d < fsilon)
        //        {
        //            pClose = pM0;
        //            index = 1;
        //            return true;
        //        }

        //        d = pM12d.DistanceTo(startP2d);
        //        if (d < fsilon)
        //        {
        //            pClose = pM1;
        //            index = 0;
        //            return true;
        //        }

        //        d = pM12d.DistanceTo(endP2d);
        //        if (d < fsilon)
        //        {
        //            pClose = pM1;
        //            index = 1;
        //            return true;
        //        }
        //    }
        //    //             else
        //    //                 return false;

        //    bool flag = false;

        //    try
        //    {
        //        var lineZ = Line.CreateBound(new XYZ(pClose2d.X, pClose2d.Y, pClose2d.Z - fsilon_1000), new XYZ(pClose2d.X, pClose2d.Y, pClose2d.Z + fsilon_1000));

        //        //Common.CreateModelLine(lineZ.GetEndPoint(0), lineZ.GetEndPoint(1));

        //        resultArray = new IntersectionResultArray();

        //        //expand
        //        double ex_ = 0.1;
        //        var curveMain_temp = Line.CreateBound(curveMain.Evaluate(-fsilon_1000, false), curveMain.Evaluate(curveMain.GetEndParameter(1) + fsilon_1000, false));

        //        //Common.CreateModelLine(curveMain.GetEndPoint(0));
        //        // Common.CreateModelLine(curveMain.GetEndPoint(1));

        //        interResult = lineZ.Intersect(curveMain_temp, out resultArray);

        //        if (interResult == SetComparisonResult.Overlap && resultArray.Size != 0)
        //        {
        //            pClose = resultArray.get_Item(0).XYZPoint;

        //            //Check
        //            var d1 = new XYZ(pClose.X, pClose.Y, 0).DistanceTo(new XYZ(pM0.X, pM0.Y, 0));
        //            var d2 = new XYZ(pClose.X, pClose.Y, 0).DistanceTo(new XYZ(pM1.X, pM1.Y, 0));

        //            if (d1 > fsilon && d2 > fsilon)
        //                isTee = true;

        //            flag = true;

        //            if (pClose.DistanceTo(startP) < pClose.DistanceTo(endP))
        //                index = 0;
        //            else
        //                index = 1;
        //        }
        //        else
        //        {
        //            var p02d_temp = new XYZ(pM0.X, pM0.Y, pClose2d.Z);
        //            var p12d_temp = new XYZ(pM1.X, pM1.Y, pClose2d.Z);

        //            var d1 = pClose2d.DistanceTo(p02d_temp);
        //            var d2 = pClose2d.DistanceTo(p12d_temp);

        //            if (d1 < fsilon)
        //            {
        //                pClose = pM0;
        //            }
        //            else if (d2 < fsilon)
        //            {
        //                pClose = pM1;
        //            }
        //            else
        //                return false;

        //            //Check
        //            if (pClose.DistanceTo(startP) < pClose.DistanceTo(endP))
        //                index = 0;
        //            else
        //                index = 1;

        //            isTee = false;
        //        }
        //    }
        //    catch (System.Exception ex)
        //    {
        //        return false;
        //    }
        //    return flag;
        //}

        //public static void gr(Pipe pipeMain, Pipe pipeCurrent, Curve pickCurve, ElementId elementTypeId, ElementId systemTypeId, ElementId levelId)
        //{
        //    var curveMain = pipeMain.GetCurve();

        //    var pM0 = curveMain.GetEndPoint(0);
        //    var pM1 = curveMain.GetEndPoint(1);

        //    var startP = pickCurve.GetEndPoint(0);
        //    var endP = pickCurve.GetEndPoint(1);

        //    int index = -1;
        //    bool isTee = false;
        //    XYZ pClose = null;
        //    if (ge(curveMain, pickCurve, out pClose, out index, out isTee) == false)
        //    {
        //        //                 var startP_temp = new XYZ(startP.X, startP.Y, pM0.Z);
        //        //                 var endP_temp = new XYZ(endP.X, endP.Y, pM0.Z);
        //        //
        //        //                 var d1 = pM0.DistanceTo(startP_temp);
        //        //                 var d3 = pM1.DistanceTo(startP_temp);
        //        //
        //        //                 var d2 = pM0.DistanceTo(endP_temp);
        //        //                 var d4 = pM1.DistanceTo(endP_temp);
        //        //
        //        //                 double fsilon = 0.001;
        //        //                 if (d1 < fsilon)
        //        //                 {
        //        //                     pClose = pM0;
        //        //                     index = 0;
        //        //                 }
        //        //                 else if (d2 < fsilon)
        //        //                 {
        //        //                     pClose = pM0;
        //        //                     index = 1;
        //        //                 }
        //        //                 else if (d3 < fsilon)
        //        //                 {
        //        //                     pClose = pM1;
        //        //                     index = 0;
        //        //                 }
        //        //                 else if (d4 < fsilon)
        //        //                 {
        //        //                     pClose = pM1;
        //        //                     index = 1;
        //        //                 }
        //        //                 isTee = false;
        //    }

        //    if (pClose == null || index == -1)
        //        return;

        //    //Find elevation

        //    if (isTee == true)
        //    {
        //        try
        //        {
        //            if (App.m_CreateMEPForm.IsCreateFitting_Tee)
        //            {
        //                //////////////////////////////////////////////////////////////////////////
        //                //                         var fitting = CreatFitting(pipeMain, pClose, pipeMain.Diameter, pipeCurrent.Diameter);
        //                //
        //                //                         ValidateFitting(newPipe, pipeCurrent, pClose, fitting);
        //                //
        //                //                         var list = new List<Pipe>() { pipeMain, newPipe, pipeCurrent };
        //                //                         Common.ConnectPipesWithTeeFitting(list, fitting, pClose);
        //                //////////////////////////////////////////////////////////////////////////

        //                CreateFittingForMEPUtils.offs(pipeMain, pipeCurrent, pClose);
        //            }
        //        }
        //        catch (System.Exception ex)
        //        {
        //        }
        //    }
        //    else
        //    {
        //        SubTransaction sub = new SubTransaction(Global.m_uiDoc.Document);
        //        sub.Start();

        //        var existCurve = (pipeMain.Location as LocationCurve).Curve;

        //        XYZ start = pClose;
        //        XYZ end = null;

        //        if (index == 1)
        //        {
        //            end = pickCurve.GetEndPoint(0);
        //        }
        //        else
        //        {
        //            end = pickCurve.GetEndPoint(1);
        //        }

        //        //Set elevation
        //        end = new XYZ(end.X, end.Y, start.Z);

        //        var z = (start - end).GetLength() * App.m_CreateMEPForm.SlopeValue_ / 100;

        //        var endPoint = new XYZ(end.X, end.Y, end.Z + (App.m_CreateMEPForm.Slope_ == Slope.Down ? z * -1 : z));

        //        //Create new mep
        //        var newLine = Line.CreateBound(start, endPoint);
        //        (pipeCurrent.Location as LocationCurve).Curve = newLine;

        //        sub.Commit();
        //        Global.m_uiDoc.Document.Regenerate();
        //        try
        //        {
        //            //Create elbow
        //            Connector c1 = Utils.GetConnectorClosestTo(
        //            pipeCurrent, pClose);

        //            Connector c2 = Utils.GetConnectorClosestTo(
        //              pipeMain, pClose);

        //            var curve_main = (pipeMain.Location as LocationCurve).Curve;
        //            var curve_current = (pipeCurrent.Location as LocationCurve).Curve;

        //            var center_main = curve_main.Evaluate(curve_main.GetEndParameter(0) + curve_main.GetEndParameter(1) / 2, false);
        //            var center_current = curve_current.Evaluate(curve_current.GetEndParameter(0) + curve_current.GetEndParameter(1) / 2, false);

        //            //Check
        //            var v1 = (Common.ToPoint2D(pClose) - Common.ToPoint2D(center_main)).Normalize();
        //            var v2 = (Common.ToPoint2D(pClose) - Common.ToPoint2D(center_current)).Normalize();

        //            var isPara = Utils.IsParallel(v1, v2, Common.fsilon);
        //            if (isPara && App.m_CreateMEPForm.IsCreateFitting_Transition)
        //            {
        //                var same = CreateFittingForMEPUtils.iss(pipeMain, pipeCurrent);
        //                if (same == true)
        //                {
        //                    //Noi lai thanh 1 pipe
        //                    CreateFittingForMEPUtils.dd(pipeMain, pipeCurrent, true);
        //                }
        //                else
        //                    Global.m_uiDoc.Document.Create.NewTransitionFitting(c1, c2);

        //                //                         if (App.m_CreateMEPForm.IsCreateFitting_Transition)
        //                //                         {
        //                //                             Global.m_uiDoc.Document.Create.NewTransitionFitting(c1, c2);
        //                //                         }
        //            }
        //            else
        //            {
        //                if (App.m_CreateMEPForm.IsCreateFitting_Elbow)
        //                {
        //                    Global.m_uiDoc.Document.Create.NewElbowFitting(c1, c2);
        //                }
        //            }
        //        }
        //        catch (System.Exception ex)
        //        {
        //        }
        //    }
        //}

        //public static MEPCurve gg(ElementId ognoreId, Curve curve, ElementId elementTypeId, ElementId systemTypeId, double width, BuiltInCategory builtInCategory, bool isCheckReducer)
        //{
        //    try
        //    {
        //        XYZ start = curve.GetEndPoint(0);
        //        XYZ end = curve.GetEndPoint(1);

        //        Element elementOut = null;
        //        double distanceOut = 0;

        //        Element minId = null;
        //        double minDis = double.MaxValue;

        //        try
        //        {
        //            var refers = Common.FindElementMEPCurves(Global.m_uiDoc.Document, ognoreId, elementTypeId, systemTypeId, start, builtInCategory, width, isCheckReducer, true, out elementOut, out distanceOut);
        //            if (refers == true && elementOut != null)
        //            {
        //                if (distanceOut == 0)
        //                    return elementOut as MEPCurve;
        //                else
        //                {
        //                    if (minDis > distanceOut)
        //                    {
        //                        minId = elementOut;
        //                        minDis = distanceOut;
        //                    }
        //                }
        //            }

        //            refers = Common.FindElementMEPCurves(Global.m_uiDoc.Document, ognoreId, elementTypeId, systemTypeId, start, builtInCategory, width, false, isCheckReducer, out elementOut, out distanceOut);
        //            if (refers == true && elementOut != null)
        //            {
        //                if (distanceOut == 0)
        //                    return elementOut as MEPCurve;
        //                else
        //                {
        //                    if (minDis > distanceOut)
        //                    {
        //                        minId = elementOut;
        //                        minDis = distanceOut;
        //                    }
        //                }
        //            }

        //            refers = Common.FindElementMEPCurves(Global.m_uiDoc.Document, ognoreId, elementTypeId, systemTypeId, end, builtInCategory, width, true, isCheckReducer, out elementOut, out distanceOut);
        //            if (refers == true && elementOut != null)
        //            {
        //                if (distanceOut == 0)
        //                    return elementOut as MEPCurve;
        //                else
        //                {
        //                    if (minDis > distanceOut)
        //                    {
        //                        minId = elementOut;
        //                        minDis = distanceOut;
        //                    }
        //                }
        //            }

        //            refers = Common.FindElementMEPCurves(Global.m_uiDoc.Document, ognoreId, elementTypeId, systemTypeId, end, builtInCategory, width, false, isCheckReducer, out elementOut, out distanceOut);
        //            if (refers == true && elementOut != null)
        //            {
        //                if (distanceOut == 0)
        //                    return elementOut as MEPCurve;
        //                else
        //                {
        //                    if (minDis > distanceOut)
        //                    {
        //                        minId = elementOut;
        //                        minDis = distanceOut;
        //                    }
        //                }
        //            }

        //            if (minId == null)
        //                return null;

        //            return minId as MEPCurve;
        //        }
        //        catch (System.Exception ex)
        //        {
        //        }

        //        return null;
        //    }
        //    catch (System.Exception ex)
        //    {
        //        return null;
        //    }
        //}

        //public static MEPCurve err(XYZ start, XYZ end, ElementId elementTypeId, ElementId systemTypeId, ElementId levelId)
        //{
        //    ElementType elementType = Global.UIDoc.Document.GetElement(elementTypeId) as ElementType;

        //    MEPCurve mepCurve = null;

        //    if (elementType is DuctType)
        //    {
        //        mepCurve = Duct.Create(Global.UIDoc.Document, systemTypeId, elementTypeId, levelId, start, end);
        //    }
        //    else if (elementType is PipeType)
        //    {
        //        mepCurve = Pipe.Create(Global.UIDoc.Document, systemTypeId, elementTypeId, levelId, start, end);
        //    }
        //    else if (elementType is CableTrayType)
        //    {
        //        mepCurve = CableTray.Create(Global.UIDoc.Document, elementTypeId, start, end, levelId);
        //    }
        //    else if (elementType is ConduitType)
        //    {
        //        mepCurve = Conduit.Create(Global.UIDoc.Document, elementTypeId, start, end, levelId);
        //    }
        //    return mepCurve;
        //}
    }

    public class err
    {
        public int i = 0;
        public List<ElementId> ei = new List<ElementId>();
        public bool ok = true;

        public err(int index)
        {
            i = index;
        }
    }
}