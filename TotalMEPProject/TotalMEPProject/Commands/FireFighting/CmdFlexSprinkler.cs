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
    public class CmdFlexSprinkler : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Global.UIApp = commandData.Application;
            Global.RVTApp = commandData.Application.Application;
            Global.UIDoc = commandData.Application.ActiveUIDocument;
            Global.AppCreation = commandData.Application.Application.Create;

            //Show form
            if (App.ShowFlexSprinklerForm() == false)
                return Result.Cancelled;

            return Result.Succeeded;
        }

        public static Result Process()
        {
            try
            {
                if (App.m_flexSprinklerForm != null && App.m_flexSprinklerForm.IsDisposed == false)
                {
                    App.m_flexSprinklerForm.Hide();
                }

                if (App.m_flexSprinklerForm.IsCheckedType1)
                {
                    ProcessType1();
                }
                else if (App.m_flexSprinklerForm.IsCheckedType2)
                {
                    ProcessType2();
                }
                else if (App.m_flexSprinklerForm.IsCheckedType3)
                {
                    ProcessType3();
                }
                return Result.Succeeded;
            }
            catch (Exception)
            { }
            finally
            {
                if (App.m_flexSprinklerForm != null && App.m_flexSprinklerForm.IsDisposed == false)
                {
                    App.m_flexSprinklerForm.Show(App.hWndRevit);
                }
                DisplayService.SetFocus(new HandleRef(null, App.m_flexSprinklerForm.Handle));
            }
            return Result.Cancelled;
        }

        public static bool ProcessType1()
        {
            try
            {
                List<FamilyInstance> sprinklers = sr.SelectSprinklers(App.m_flexSprinklerForm.IsCheckedD15);
                if (sprinklers == null || sprinklers.Count == 0)
                    return false;

                List<Pipe> pipes = SelectPipes();
                if (pipes == null || pipes.Count == 0)
                    return false;

                var pipeIds = (from Pipe p in pipes
                               where p.Id != ElementId.InvalidElementId
                               select p.Id).ToList();

                var height = App.m_flexSprinklerForm.VerticalPipeLengthL2;

                Transaction tran = new Transaction(Global.UIDoc.Document, "CreateConnector");
                tran.Start();

                double radius = 500/*400*/; //mm : sua thanh 500 theo yeu cau cua a Cuong
                var ft = Common.mmToFT * radius;

                // Status cancel export : default = false
                bool isCancelExport = false;

                // Count type imported
                int nCount = 0;

                // Initialize progress bar
                System.Diagnostics.Process process = System.Diagnostics.Process.GetCurrentProcess();
                IntPtr intPtr = process.MainWindowHandle;
                ViewSingleProgressBar progressBar = new ViewSingleProgressBar("Flex Sprinkler", "Process : ");
                progressBar.prgSingle.Minimum = 1;
                progressBar.prgSingle.Maximum = sprinklers.Count;
                progressBar.prgSingle.Value = 1;
                WindowInteropHelper helper = new WindowInteropHelper(progressBar);
                helper.Owner = intPtr;
                progressBar.Show();

                try
                {
                    foreach (FamilyInstance instance in sprinklers)
                    {
                        double dPercent = 0.0;
                        var sprinkle_point = (instance.Location as LocationPoint).Point;

                        //Check connect
                        var connects = instance.MEPModel.ConnectorManager.Connectors;

                        if (connects.Size == 0)
                        {
                            nCount++;
                            dPercent = (nCount / (sprinklers.Count * 1.0)) * 100.0;
                            progressBar.tbxMessage.Text = "Complete : " + dPercent.ToString("0.00") + "% ";
                            progressBar.IncrementProgressBar();
                            continue;
                        }

                        var connect = instance.MEPModel.ConnectorManager.Lookup(1);

                        if (connect.IsConnected == true)
                        {
                            nCount++;
                            dPercent = (nCount / (sprinklers.Count * 1.0)) * 100.0;
                            progressBar.tbxMessage.Text = "Complete : " + dPercent.ToString("0.00") + "% ";
                            progressBar.IncrementProgressBar();
                            continue;
                        }

                        var solid = Common.CreateCylindricalVolume(sprinkle_point, ft * 5, ft, true);
                        if (solid == null)
                        {
                            nCount++;
                            dPercent = (nCount / (sprinklers.Count * 1.0)) * 100.0;
                            progressBar.tbxMessage.Text = "Complete : " + dPercent.ToString("0.00") + "% ";
                            progressBar.IncrementProgressBar();
                            continue;
                        }

                        //Find intersection

                        FilteredElementCollector collector = new FilteredElementCollector(Global.UIDoc.Document, pipeIds);
                        collector.OfClass(typeof(Pipe));
                        collector.WherePasses(new ElementIntersectsSolidFilter(solid)); // Apply intersection filter to find matches
                        if (collector.GetElementCount() == 0)
                            continue;
                        var pipeList = collector.ToElements();

                        //Global.m_uiDoc.Selection.SetElementIds(collector.ToElementIds());

                        bool split = false;
                        var pipe = ProcessPipes(pipeList.ToList(), sprinkle_point, out split);

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

                        bool isDauOng = false;
                        IntersectionResultArray arr = new IntersectionResultArray();
                        if (split == false)
                        {
                            //truong hop o dau ong

                            //expand
                            var index = line2d.GetEndPoint(0).DistanceTo(sprinkle_point) < line2d.GetEndPoint(1).DistanceTo(sprinkle_point) ? 0 : 1;
                            var curveExpand = Line.CreateUnbound(line2d.GetEndPoint(index), line2d.Direction * 100);

                            var inter = curveExpand.Intersect(lineTemp, out arr);
                            if (inter != SetComparisonResult.Overlap)
                            {
                                nCount++;
                                dPercent = (nCount / (sprinklers.Count * 1.0)) * 100.0;
                                progressBar.tbxMessage.Text = "Complete : " + dPercent.ToString("0.00") + "% ";
                                progressBar.IncrementProgressBar();
                                continue;
                            }

                            var p2d = arr.get_Item(0).XYZPoint;

                            var p3d = new XYZ(p2d.X, p2d.Y, sprinkle_point.Z);

                            var line3d = Line.CreateBound(p3d, new XYZ(p3d.X, p3d.Y, p3d.Z + dTemp));

                            var curveExtend3d_temp = Line.CreateUnbound(curve.GetEndPoint(index), (curve as Line).Direction * 100);

                            arr = new IntersectionResultArray();
                            inter = curveExtend3d_temp.Intersect(line3d, out arr);
                            if (inter != SetComparisonResult.Overlap)
                            {
                                nCount++;
                                dPercent = (nCount / (sprinklers.Count * 1.0)) * 100.0;
                                progressBar.tbxMessage.Text = "Complete : " + dPercent.ToString("0.00") + "% ";
                                progressBar.IncrementProgressBar();
                                continue;
                            }

                            p = arr.get_Item(0).XYZPoint;
                        }
                        else
                        {
                            var inter = lineTemp.Intersect(line2d, out arr);
                            if (inter != SetComparisonResult.Overlap)
                            {
                                nCount++;
                                dPercent = (nCount / (sprinklers.Count * 1.0)) * 100.0;
                                progressBar.tbxMessage.Text = "Complete : " + dPercent.ToString("0.00") + "% ";
                                progressBar.IncrementProgressBar();
                                continue;
                            }

                            var p2d = arr.get_Item(0).XYZPoint;

                            var p3d = new XYZ(p2d.X, p2d.Y, sprinkle_point.Z);

                            var line3d = Line.CreateBound(p3d, new XYZ(p3d.X, p3d.Y, p3d.Z + dTemp));

                            arr = new IntersectionResultArray();
                            inter = curve.Intersect(line3d, out arr);
                            if (inter != SetComparisonResult.Overlap)
                            {
                                nCount++;
                                dPercent = (nCount / (sprinklers.Count * 1.0)) * 100.0;
                                progressBar.tbxMessage.Text = "Complete : " + dPercent.ToString("0.00") + "% ";
                                progressBar.IncrementProgressBar();
                                continue;
                            }

                            p = arr.get_Item(0).XYZPoint;

                            pipe2 = null;
                            bool flagCreateTee = true;
                            if (GetPreferredJunctionType(pipe) != PreferredJunctionType.Tee)
                            {
                                flagCreateTee = false;
                            }

                            ProcessStartSidePipe(pipe, out pipe2, p, out isDauOng, flagCreateTee);

                            if (pipe2 != null)
                            {
                                pipeIds.Add(pipe2.Id);
                            }
                        }

                        //Set d = 25

                        double dFt = App.m_flexSprinklerForm.PipeSize * Common.mmToFT;

                        var ft_h = Common.mmToFT * height;

                        newPlace = new XYZ(0, 0, 0);
                        elemIds = ElementTransformUtils.CopyElement(
                         Global.UIDoc.Document, pipe1.Id, newPlace);

                        var pipe_v1 = Global.UIDoc.Document.GetElement(elemIds.ToList()[0]) as Pipe;

                        var line_v1 = Line.CreateUnbound(p, XYZ.BasisZ * ft_h * 2);

                        line_v1 = Line.CreateBound(p, line_v1.Evaluate(ft_h, false));
                        (pipe_v1.Location as LocationCurve).Curve = line_v1;

                        pipe_v1.LookupParameter("Diameter").Set(dFt);

                        var pipeType = Global.UIDoc.Document.GetElement(App.m_flexSprinklerForm.FamilyType) as PipeType;
                        if (pipeType != null)
                            pipe_v1.PipeType = pipeType;
                        //Connect
                        try
                        {
                            var c1 = Common.GetConnectorClosestTo(pipe1, p);
                            var c3 = Common.GetConnectorClosestTo(pipe_v1, p);

                            if (App.m_flexSprinklerForm.IsCheckedTee)
                            {
                                if (GetPreferredJunctionType(pipe1) != PreferredJunctionType.Tee)
                                {
                                    CreateTap(pipe1 as MEPCurve, pipe_v1 as MEPCurve);
                                }
                                else
                                {
                                    if (pipe2 != null)
                                    {
                                        var c2 = Common.GetConnectorClosestTo(pipe2, p);
                                        var fitting = Global.UIDoc.Document.Create.NewTeeFitting(c1, c2, c3);
                                    }
                                    else
                                    {
                                        tran.RollBack();
                                        continue;
                                    }
                                }
                            }
                            else
                            {
                                if (!isDauOng)
                                {
                                    if (GetPreferredJunctionType(pipe1) != PreferredJunctionType.Tee)
                                    {
                                        CreateTap(pipe1 as MEPCurve, pipe_v1 as MEPCurve);
                                    }
                                    else
                                    {
                                        if (pipe2 != null)
                                        {
                                            var c2 = Common.GetConnectorClosestTo(pipe2, p);
                                            var fitting = Global.UIDoc.Document.Create.NewTeeFitting(c1, c2, c3);
                                        }
                                    }
                                }
                                else
                                {
                                    Global.UIDoc.Document.Create.NewElbowFitting(c1, c3);
                                }
                            }
                        }
                        catch (System.Exception ex)
                        {
                        }

                        //Hor
                        var v_v = (new XYZ(sprinkle_point.X, sprinkle_point.Y, 0) - new XYZ(p.X, p.Y, 0)).Normalize();

                        var ft_v = App.m_flexSprinklerForm.HorizontalPipeLengthL * Common.mmToFT;

                        var line_Extend = Line.CreateUnbound(line_v1.GetEndPoint(1), ft_v * v_v * 2);

                        newPlace = new XYZ(0, 0, 0);
                        elemIds = ElementTransformUtils.CopyElement(
                         Global.UIDoc.Document, pipe1.Id, newPlace);

                        var pipe_hor = Global.UIDoc.Document.GetElement(elemIds.ToList()[0]) as Pipe;
                        var line_hor = Line.CreateBound(line_v1.GetEndPoint(1), line_Extend.Evaluate(ft_v, false));

                        (pipe_hor.Location as LocationCurve).Curve = line_hor;

                        pipe_hor.LookupParameter("Diameter").Set(dFt);

                        //Vertical 2
                        var line_v2 = Line.CreateBound(line_hor.GetEndPoint(1), sprinkle_point);
                        var center = line_v2.Evaluate((line_v2.GetEndParameter(0) + line_v2.GetEndParameter(1)) / 2, false);

                        try
                        {
                            var c1 = Common.GetConnectorClosestTo(pipe_hor, line_hor.GetEndPoint(1));
                            var c2 = Common.GetConnectorClosestTo(instance, center);

                            FlexPipeType flexPipeType = new FilteredElementCollector(Global.UIDoc.Document).OfCategory(BuiltInCategory.OST_FlexPipeCurves).OfClass(typeof(FlexPipeType)).WhereElementIsElementType().Cast<FlexPipeType>().FirstOrDefault();

                            List<XYZ> pnts = new List<XYZ>();

                            pnts.Add(c1.Origin);
                            pnts.Add(c2.Origin);

                            var flexPipe = Global.UIDoc.Document.Create.NewFlexPipe(pnts, flexPipeType);

                            //Set the diameter of flex.
                            flexPipe.LookupParameter("Diameter").Set(dFt);
                            flexPipe.StartTangent = c1.CoordinateSystem.BasisZ; //Set Tangent to Conn Direction.
                            flexPipe.EndTangent = c2.CoordinateSystem.BasisZ.Negate();

                            var c1_flexPipe = GetConnectorClosestTo(flexPipe, c1.Origin);
                            var c2_flexPipe = GetConnectorClosestTo(flexPipe, c2.Origin);

                            var union = Global.UIDoc.Document.Create.NewUnionFitting(c1_flexPipe, c1);
                            var con = Common.ToList(union.MEPModel.ConnectorManager.Connectors);

                            var vectorMove = con[1].Origin - con[0].Origin;

                            ElementTransformUtils.MoveElement(Global.UIDoc.Document, union.Id, vectorMove.Normalize() * 0.00001);

                            var paraTransition = flexPipeType.get_Parameter(BuiltInParameter.RBS_CURVETYPE_DEFAULT_TRANSITION_PARAM);
                            if (paraTransition != null && !paraTransition.IsReadOnly)
                            {
                                var fmlReducerSymbol = new FilteredElementCollector(Global.UIDoc.Document).OfClass(typeof(FamilySymbol))
                                                                                                          .Cast<FamilySymbol>().FirstOrDefault(x => x.FamilyName.Contains("Reducer - Threaded - GSP"));
                                if (fmlReducerSymbol != null)
                                    paraTransition.Set(fmlReducerSymbol.Id);
                                Global.UIDoc.Document.Create.NewTransitionFitting(c2_flexPipe, c2);
                            }

                            Global.UIDoc.Document.Regenerate();
                        }
                        catch (Exception)
                        { }

                        try
                        {
                            var c1 = Common.GetConnectorClosestTo(pipe_v1, line_v1.GetEndPoint(1));
                            var c2 = Common.GetConnectorClosestTo(pipe_hor, line_v1.GetEndPoint(1));

                            Global.UIDoc.Document.Create.NewElbowFitting(c1, c2);
                        }
                        catch (System.Exception ex)
                        {
                        }

                        nCount++;
                        dPercent = (nCount / (sprinklers.Count * 1.0)) * 100.0;
                        progressBar.tbxMessage.Text = "Complete : " + dPercent.ToString("0.00") + "% ";
                        progressBar.IncrementProgressBar();
                    }

                    if (isCancelExport == false)
                        progressBar.Dispose();
                }
                catch (Exception)
                { }
                finally
                { progressBar.Dispose(); }
                //Find pipe

                tran.Commit();
            }
            catch (Exception)
            { }
            return false;
        }

        public static bool ProcessType2()
        {
            try
            {
                List<FamilyInstance> sprinklers = sr.SelectSprinklers(App.m_flexSprinklerForm.IsCheckedD15);
                if (sprinklers == null || sprinklers.Count == 0)
                    return false;

                List<Pipe> pipes = SelectPipes();
                if (pipes == null || pipes.Count == 0)
                    return false;

                var pipeIds = (from Pipe p in pipes
                               where p.Id != ElementId.InvalidElementId
                               select p.Id).ToList();

                var height = App.m_flexSprinklerForm.VerticalPipeLengthL2;

                Transaction tran = new Transaction(Global.UIDoc.Document, "CreateConnector");
                tran.Start();

                double radius = 500/*400*/; //mm : sua thanh 500 theo yeu cau cua a Cuong
                var ft = Common.mmToFT * radius;

                // Status cancel export : default = false
                bool isCancelExport = false;

                // Count type imported
                int nCount = 0;

                // Initialize progress bar
                System.Diagnostics.Process process = System.Diagnostics.Process.GetCurrentProcess();
                IntPtr intPtr = process.MainWindowHandle;
                ViewSingleProgressBar progressBar = new ViewSingleProgressBar("Flex Sprinkler", "Process : ");
                progressBar.prgSingle.Minimum = 1;
                progressBar.prgSingle.Maximum = sprinklers.Count;
                progressBar.prgSingle.Value = 1;
                WindowInteropHelper helper = new WindowInteropHelper(progressBar);
                helper.Owner = intPtr;
                progressBar.Show();

                try
                {
                    //Find pipe
                    foreach (FamilyInstance instance in sprinklers)
                    {
                        double dPercent = 0.0;
                        var sprinkle_point = (instance.Location as LocationPoint).Point;

                        //Check connect
                        var connects = instance.MEPModel.ConnectorManager.Connectors;

                        if (connects.Size == 0)
                        {
                            nCount++;
                            dPercent = (nCount / (sprinklers.Count * 1.0)) * 100.0;
                            progressBar.tbxMessage.Text = "Complete : " + dPercent.ToString("0.00") + "% ";
                            progressBar.IncrementProgressBar();
                            continue;
                        }

                        var connect = instance.MEPModel.ConnectorManager.Lookup(1);

                        if (connect.IsConnected == true)
                        {
                            nCount++;
                            dPercent = (nCount / (sprinklers.Count * 1.0)) * 100.0;
                            progressBar.tbxMessage.Text = "Complete : " + dPercent.ToString("0.00") + "% ";
                            progressBar.IncrementProgressBar();
                            continue;
                        }

                        var solid = Common.CreateCylindricalVolume(sprinkle_point, ft * 5, ft, true);
                        if (solid == null)
                        {
                            nCount++;
                            dPercent = (nCount / (sprinklers.Count * 1.0)) * 100.0;
                            progressBar.tbxMessage.Text = "Complete : " + dPercent.ToString("0.00") + "% ";
                            progressBar.IncrementProgressBar();
                            continue;
                        }

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
                        var pipe = ProcessPipes(pipeList.ToList(), sprinkle_point, out split);

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

                        bool isDauOng = false;
                        IntersectionResultArray arr = new IntersectionResultArray();
                        if (split == false)
                        {
                            //truong hop o dau ong

                            //expand
                            var index = line2d.GetEndPoint(0).DistanceTo(sprinkle_point) < line2d.GetEndPoint(1).DistanceTo(sprinkle_point) ? 0 : 1;
                            var curveExpand = Line.CreateUnbound(line2d.GetEndPoint(index), line2d.Direction * 100);

                            var inter = curveExpand.Intersect(lineTemp, out arr);
                            if (inter != SetComparisonResult.Overlap)
                            {
                                nCount++;
                                dPercent = (nCount / (sprinklers.Count * 1.0)) * 100.0;
                                progressBar.tbxMessage.Text = "Complete : " + dPercent.ToString("0.00") + "% ";
                                progressBar.IncrementProgressBar();
                                continue;
                            }

                            var p2d = arr.get_Item(0).XYZPoint;

                            var p3d = new XYZ(p2d.X, p2d.Y, sprinkle_point.Z);

                            var line3d = Line.CreateBound(p3d, new XYZ(p3d.X, p3d.Y, p3d.Z + dTemp));

                            var curveExtend3d_temp = Line.CreateUnbound(curve.GetEndPoint(index), (curve as Line).Direction * 100);

                            arr = new IntersectionResultArray();
                            inter = curveExtend3d_temp.Intersect(line3d, out arr);
                            if (inter != SetComparisonResult.Overlap)
                            {
                                nCount++;
                                dPercent = (nCount / (sprinklers.Count * 1.0)) * 100.0;
                                progressBar.tbxMessage.Text = "Complete : " + dPercent.ToString("0.00") + "% ";
                                progressBar.IncrementProgressBar();
                                continue;
                            }

                            p = arr.get_Item(0).XYZPoint;
                        }
                        else
                        {
                            var inter = lineTemp.Intersect(line2d, out arr);
                            if (inter != SetComparisonResult.Overlap)
                            {
                                nCount++;
                                dPercent = (nCount / (sprinklers.Count * 1.0)) * 100.0;
                                progressBar.tbxMessage.Text = "Complete : " + dPercent.ToString("0.00") + "% ";
                                progressBar.IncrementProgressBar();
                                continue;
                            }

                            var p2d = arr.get_Item(0).XYZPoint;

                            var p3d = new XYZ(p2d.X, p2d.Y, sprinkle_point.Z);

                            var line3d = Line.CreateBound(p3d, new XYZ(p3d.X, p3d.Y, p3d.Z + dTemp));

                            arr = new IntersectionResultArray();
                            inter = curve.Intersect(line3d, out arr);
                            if (inter != SetComparisonResult.Overlap)
                            {
                                nCount++;
                                dPercent = (nCount / (sprinklers.Count * 1.0)) * 100.0;
                                progressBar.tbxMessage.Text = "Complete : " + dPercent.ToString("0.00") + "% ";
                                progressBar.IncrementProgressBar();
                                continue;
                            }

                            p = arr.get_Item(0).XYZPoint;

                            pipe2 = null;
                            bool flagCreateTee = true;
                            if (GetPreferredJunctionType(pipe) != PreferredJunctionType.Tee)
                            {
                                flagCreateTee = false;
                            }

                            ProcessStartSidePipe(pipe, out pipe2, p, out isDauOng, flagCreateTee);

                            if (pipe2 != null)
                            {
                                pipeIds.Add(pipe2.Id);
                            }
                        }

                        //Set d = 25

                        var dFt = App.m_flexSprinklerForm.PipeSize * Common.mmToFT;

                        var ft_h = Common.mmToFT * height;

                        newPlace = new XYZ(0, 0, 0);
                        elemIds = ElementTransformUtils.CopyElement(
                         Global.UIDoc.Document, pipe1.Id, newPlace);

                        var pipe_v1 = Global.UIDoc.Document.GetElement(elemIds.ToList()[0]) as Pipe;

                        var line_v1 = Line.CreateUnbound(p, XYZ.BasisZ * ft_h * 2);

                        line_v1 = Line.CreateBound(p, line_v1.Evaluate(ft_h, false));
                        (pipe_v1.Location as LocationCurve).Curve = line_v1;

                        pipe_v1.LookupParameter("Diameter").Set(dFt);
                        var pipeType = Global.UIDoc.Document.GetElement(App.m_flexSprinklerForm.FamilyType) as PipeType;
                        if (pipeType != null)
                            pipe_v1.PipeType = pipeType;
                        //Connect
                        try
                        {
                            var c1 = Common.GetConnectorClosestTo(pipe1, p);
                            var c3 = Common.GetConnectorClosestTo(pipe_v1, p);

                            if (App.m_flexSprinklerForm.IsCheckedTee)
                            {
                                if (GetPreferredJunctionType(pipe1) != PreferredJunctionType.Tee)
                                {
                                    CreateTap(pipe1 as MEPCurve, pipe_v1 as MEPCurve);
                                }
                                else
                                {
                                    if (pipe2 != null)
                                    {
                                        var c2 = Common.GetConnectorClosestTo(pipe2, p);
                                        var fitting = Global.UIDoc.Document.Create.NewTeeFitting(c1, c2, c3);
                                    }
                                    else
                                    {
                                        tran.RollBack();
                                        continue;
                                    }
                                }
                            }
                            else
                            {
                                if (!isDauOng)
                                {
                                    if (GetPreferredJunctionType(pipe1) != PreferredJunctionType.Tee)
                                    {
                                        CreateTap(pipe1 as MEPCurve, pipe_v1 as MEPCurve);
                                    }
                                    else
                                    {
                                        if (pipe2 != null)
                                        {
                                            var c2 = Common.GetConnectorClosestTo(pipe2, p);
                                            var fitting = Global.UIDoc.Document.Create.NewTeeFitting(c1, c2, c3);
                                        }
                                    }
                                }
                                else
                                {
                                    Global.UIDoc.Document.Create.NewElbowFitting(c1, c3);
                                }
                            }
                        }
                        catch (System.Exception ex)
                        {
                        }

                        //Hor
                        var v_v = (new XYZ(sprinkle_point.X, sprinkle_point.Y, 0) - new XYZ(p.X, p.Y, 0)).Normalize();

                        var ft_v = App.m_flexSprinklerForm.HorizontalPipeLengthL * Common.mmToFT;

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

                        XYZ directionTemp = (sprinkle_point - line_hor.GetEndPoint(1)).Normalize();
                        var line_v2 = Line.CreateBound(line_hor.GetEndPoint(1), line_hor.GetEndPoint(1) + XYZ.BasisZ.Negate() * App.m_flexSprinklerForm.ExtendPipeLengthL1 * Common.mmToFT);

                        newPlace = new XYZ(0, 0, 0);
                        elemIds = ElementTransformUtils.CopyElement(
                         Global.UIDoc.Document, pipe1.Id, newPlace);

                        var pipe_v2 = Global.UIDoc.Document.GetElement(elemIds.ToList()[0]) as Pipe;
                        var center = line_hor.GetEndPoint(1) + directionTemp * App.m_flexSprinklerForm.ExtendPipeLengthL1 * Common.mmToFT;

                        (pipe_v2.Location as LocationCurve).Curve = line_v2;

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

                            FlexPipeType flexPipeType = new FilteredElementCollector(Global.UIDoc.Document).OfCategory(BuiltInCategory.OST_FlexPipeCurves).OfClass(typeof(FlexPipeType)).WhereElementIsElementType().Cast<FlexPipeType>().FirstOrDefault();

                            List<XYZ> pnts = new List<XYZ>();

                            pnts.Add(c1.Origin);
                            pnts.Add(c2.Origin);

                            var flexPipe = Global.UIDoc.Document.Create.NewFlexPipe(pnts, flexPipeType);

                            //Set the diameter of flex.
                            flexPipe.LookupParameter("Diameter").Set(dFt);
                            flexPipe.StartTangent = c1.CoordinateSystem.BasisZ; //Set Tangent to Conn Direction.
                            flexPipe.EndTangent = c2.CoordinateSystem.BasisZ.Negate();

                            var c1_flexPipe = GetConnectorClosestTo(flexPipe, c1.Origin);
                            var c2_flexPipe = GetConnectorClosestTo(flexPipe, c2.Origin);

                            var union = Global.UIDoc.Document.Create.NewUnionFitting(c1_flexPipe, c1);
                            var con = Common.ToList(union.MEPModel.ConnectorManager.Connectors);

                            var vectorMove = con[1].Origin - con[0].Origin;

                            ElementTransformUtils.MoveElement(Global.UIDoc.Document, union.Id, vectorMove.Normalize() * 0.00001);

                            var paraTransition = flexPipeType.get_Parameter(BuiltInParameter.RBS_CURVETYPE_DEFAULT_TRANSITION_PARAM);
                            if (paraTransition != null && !paraTransition.IsReadOnly)
                            {
                                var fmlReducerSymbol = new FilteredElementCollector(Global.UIDoc.Document).OfClass(typeof(FamilySymbol))
                                                                                                          .Cast<FamilySymbol>().FirstOrDefault(x => x.FamilyName.Contains("Reducer - Threaded - GSP"));
                                if (fmlReducerSymbol != null)
                                    paraTransition.Set(fmlReducerSymbol.Id);
                                Global.UIDoc.Document.Create.NewTransitionFitting(c2_flexPipe, c2);
                            }

                            Global.UIDoc.Document.Regenerate();
                        }
                        catch (Exception)
                        { }

                        //try
                        //{
                        //    var c1 = Common.GetConnectorClosestTo(pipe_v2, center);
                        //    var c2 = Common.GetConnectorClosestTo(instance, center);

                        //    Global.UIDoc.Document.Create.NewTransitionFitting(c1, c2);
                        //}
                        //catch (System.Exception ex)
                        //{
                        //}

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

                    if (isCancelExport == false)
                        progressBar.Dispose();
                }
                catch (Exception)
                { }
                finally
                { progressBar.Dispose(); }

                tran.Commit();
            }
            catch (Exception)
            { }
            return false;
        }

        public static bool ProcessType3()
        {
            try
            {
                // Get selected sprinkler
                List<FamilyInstance> selSprinklers = sr.SelectSprinklers(App.m_flexSprinklerForm.IsCheckedD15);
                if (selSprinklers == null || selSprinklers.Count == 0)
                    return false;

                // Get selected main pipe
                List<Pipe> selPipes = SelectPipes();
                if (selPipes == null || selPipes.Count == 0)
                    return false;

                // Get selected main pipe id
                List<ElementId> selPipeIds = selPipes.Where(item => item.Id != ElementId.InvalidElementId).Select(item => item.Id).ToList();

                // Process
                Transaction reTrans = new Transaction(Global.UIDoc.Document, "SPRINKLER_DOWN_RIGHT_TYPE_3");
                reTrans.Start();

                try
                {
                    double invalidRadius_mm = 500;
                    double invalidRadius_ft = Common.mmToFT * invalidRadius_mm;

                    // Status cancel export : default = false
                    bool isCancelExport = false;

                    // Count type imported
                    int nCount = 0;

                    // Initialize progress bar
                    System.Diagnostics.Process process = System.Diagnostics.Process.GetCurrentProcess();
                    IntPtr intPtr = process.MainWindowHandle;
                    ViewSingleProgressBar progressBar = new ViewSingleProgressBar("Flex Sprinkler", "Process : ");
                    progressBar.prgSingle.Minimum = 1;
                    progressBar.prgSingle.Maximum = selSprinklers.Count;
                    progressBar.prgSingle.Value = 1;
                    WindowInteropHelper helper = new WindowInteropHelper(progressBar);
                    helper.Owner = intPtr;
                    progressBar.Show();

                    try
                    {
                        foreach (FamilyInstance sprinkler in selSprinklers)
                        {
                            double dPercent = 0.0;
                            // Location sprinkler
                            XYZ locSprinkler = (sprinkler.Location as LocationPoint).Point;

                            // Check valid connect
                            ConnectorSet cntSetOfIns = sprinkler.MEPModel.ConnectorManager.Connectors;

                            if (cntSetOfIns.Size == 0)
                            {
                                nCount++;
                                dPercent = (nCount / (selSprinklers.Count * 1.0)) * 100.0;
                                progressBar.tbxMessage.Text = "Complete : " + dPercent.ToString("0.00") + "% ";
                                progressBar.IncrementProgressBar();
                                continue;
                            }

                            Connector cntOfIns_1 = sprinkler.MEPModel.ConnectorManager.Lookup(1);

                            if (cntOfIns_1.IsConnected == true)
                            {
                                nCount++;
                                dPercent = (nCount / (selSprinklers.Count * 1.0)) * 100.0;
                                progressBar.tbxMessage.Text = "Complete : " + dPercent.ToString("0.00") + "% ";
                                progressBar.IncrementProgressBar();
                                continue;
                            }

                            // Find intersection with sprinkler
                            var cylindricalFromIns = Common.CreateCylindricalVolume(locSprinkler, invalidRadius_ft * 5, invalidRadius_ft, true);
                            if (cylindricalFromIns == null)
                            {
                                nCount++;
                                dPercent = (nCount / (selSprinklers.Count * 1.0)) * 100.0;
                                progressBar.tbxMessage.Text = "Complete : " + dPercent.ToString("0.00") + "% ";
                                progressBar.IncrementProgressBar();
                                continue;
                            }

                            FilteredElementCollector filterCollector = new FilteredElementCollector(Global.UIDoc.Document, selPipeIds).OfClass(typeof(Pipe)).WherePasses(new ElementIntersectsSolidFilter(cylindricalFromIns));
                            if (filterCollector == null || filterCollector.GetElementCount() <= 0)
                            {
                                nCount++;
                                dPercent = (nCount / (selSprinklers.Count * 1.0)) * 100.0;
                                progressBar.tbxMessage.Text = "Complete : " + dPercent.ToString("0.00") + "% ";
                                progressBar.IncrementProgressBar();
                                continue;
                            }

                            IList<Element> validPipes = filterCollector.ToElements();

                            // Check pipe nesscesary split
                            bool isSplit = false;
                            Pipe processPipe = ProcessPipes(validPipes.ToList(), locSprinkler, out isSplit);

                            // Process main pipe
                            Curve curveProcessPipe = (processPipe.Location as LocationCurve).Curve;

                            XYZ firstPnt_ProcessPipe = curveProcessPipe.GetEndPoint(0);
                            XYZ secondPnt_ProcessPipe = curveProcessPipe.GetEndPoint(1);

                            double dTempEvaluate = 1000;

                            var curveProcessPipe_2d = Line.CreateBound(new XYZ(firstPnt_ProcessPipe.X, firstPnt_ProcessPipe.Y, 0), new XYZ(secondPnt_ProcessPipe.X, secondPnt_ProcessPipe.Y, 0));
                            XYZ dirCrossProduct = curveProcessPipe_2d.Direction.CrossProduct(XYZ.BasisZ).Normalize();

                            var curveProcessPipe_crossProduct_2d = Line.CreateUnbound(new XYZ(locSprinkler.X, locSprinkler.Y, 0), dirCrossProduct * dTempEvaluate);

                            XYZ p11 = curveProcessPipe_crossProduct_2d.Evaluate(dTempEvaluate, false);

                            curveProcessPipe_crossProduct_2d = Line.CreateUnbound(new XYZ(locSprinkler.X, locSprinkler.Y, 0), -dirCrossProduct * dTempEvaluate);

                            XYZ p22 = curveProcessPipe_crossProduct_2d.Evaluate(dTempEvaluate, false);

                            curveProcessPipe_crossProduct_2d = Line.CreateBound(p11, p22);

                            // Find intersection point
                            XYZ newPlace = new XYZ(0, 0, 0);
                            ICollection<ElementId> elemIds = null;
                            var temp_processPipe_1 = processPipe;
                            Pipe temp_processPipe_2 = null;
                            XYZ finalIntPnt = null;

                            IntersectionResultArray intRetArr = new IntersectionResultArray();

                            bool isDauOng = false;
                            //Truong hop dau ong
                            if (isSplit == false)
                            {
                                //Expand
                                var index = curveProcessPipe_2d.GetEndPoint(0).DistanceTo(locSprinkler) < curveProcessPipe_2d.GetEndPoint(1).DistanceTo(locSprinkler) ? 0 : 1;
                                var curveExpand = Line.CreateUnbound(curveProcessPipe_2d.GetEndPoint(index), curveProcessPipe_2d.Direction * 100);

                                var inter = curveExpand.Intersect(curveProcessPipe_crossProduct_2d, out intRetArr);
                                if (inter != SetComparisonResult.Overlap)
                                {
                                    nCount++;
                                    dPercent = (nCount / (selSprinklers.Count * 1.0)) * 100.0;
                                    progressBar.tbxMessage.Text = "Complete : " + dPercent.ToString("0.00") + "% ";
                                    progressBar.IncrementProgressBar();
                                    continue;
                                }

                                var p2d = intRetArr.get_Item(0).XYZPoint;

                                var p3d = new XYZ(p2d.X, p2d.Y, locSprinkler.Z);

                                var line3d = Line.CreateBound(p3d, new XYZ(p3d.X, p3d.Y, p3d.Z + dTempEvaluate));
                                var curveExtend3d_temp = Line.CreateUnbound(curveProcessPipe.GetEndPoint(index), (curveProcessPipe as Line).Direction * 100);

                                intRetArr = new IntersectionResultArray();
                                inter = curveExtend3d_temp.Intersect(line3d, out intRetArr);
                                if (inter != SetComparisonResult.Overlap)
                                {
                                    nCount++;
                                    dPercent = (nCount / (selSprinklers.Count * 1.0)) * 100.0;
                                    progressBar.tbxMessage.Text = "Complete : " + dPercent.ToString("0.00") + "% ";
                                    progressBar.IncrementProgressBar();
                                    continue;
                                }

                                finalIntPnt = intRetArr.get_Item(0).XYZPoint;
                            }

                            // Truong hop o giua ong
                            else
                            {
                                var inter = curveProcessPipe_crossProduct_2d.Intersect(curveProcessPipe_2d, out intRetArr);
                                if (inter != SetComparisonResult.Overlap)
                                {
                                    nCount++;
                                    dPercent = (nCount / (selSprinklers.Count * 1.0)) * 100.0;
                                    progressBar.tbxMessage.Text = "Complete : " + dPercent.ToString("0.00") + "% ";
                                    progressBar.IncrementProgressBar();
                                    continue;
                                }

                                var p2d = intRetArr.get_Item(0).XYZPoint;

                                var p3d = new XYZ(p2d.X, p2d.Y, locSprinkler.Z);

                                var line3d = Line.CreateBound(p3d, new XYZ(p3d.X, p3d.Y, p3d.Z + dTempEvaluate));

                                intRetArr = new IntersectionResultArray();
                                inter = curveProcessPipe.Intersect(line3d, out intRetArr);
                                if (inter != SetComparisonResult.Overlap)
                                {
                                    nCount++;
                                    dPercent = (nCount / (selSprinklers.Count * 1.0)) * 100.0;
                                    progressBar.tbxMessage.Text = "Complete : " + dPercent.ToString("0.00") + "% ";
                                    progressBar.IncrementProgressBar();
                                    continue;
                                }

                                finalIntPnt = intRetArr.get_Item(0).XYZPoint;

                                temp_processPipe_2 = null;
                                bool flagCreateTee = true;
                                if (GetPreferredJunctionType(processPipe) != PreferredJunctionType.Tee)
                                {
                                    flagCreateTee = false;
                                }

                                ProcessStartSidePipe(processPipe, out temp_processPipe_2, finalIntPnt, out isDauOng, flagCreateTee);

                                if (temp_processPipe_2 != null)
                                {
                                    selPipeIds.Add(temp_processPipe_2.Id);
                                }
                            }

                            ////Set pipe size
                            //var dPipeSizeFt = (double)GetParameterValueByName(processPipe, "Diameter"); ;

                            // Generate Pipe Horizontal
                            var v_v = (new XYZ(locSprinkler.X, locSprinkler.Y, 0) - new XYZ(finalIntPnt.X, finalIntPnt.Y, 0)).Normalize();
                            var ft_v = (new XYZ(locSprinkler.X, locSprinkler.Y, 0) - new XYZ(finalIntPnt.X, finalIntPnt.Y, 0)).GetLength();
                            var line_Extend = Line.CreateUnbound(finalIntPnt, ft_v * v_v * 2);

                            newPlace = new XYZ(0, 0, 0);
                            elemIds = ElementTransformUtils.CopyElement(
                             Global.UIDoc.Document, temp_processPipe_1.Id, newPlace);

                            var horizontal_pipe = Global.UIDoc.Document.GetElement(elemIds.ToList()[0]) as Pipe;
                            var hor_line = Line.CreateBound(finalIntPnt, line_Extend.Evaluate(App.m_flexSprinklerForm.HorizontalPipeLengthL * Common.mmToFT, false));
                            (horizontal_pipe.Location as LocationCurve).Curve = hor_line;
                            horizontal_pipe.LookupParameter("Diameter").Set(App.m_flexSprinklerForm.PipeSize * Common.mmToFT);

                            // Connect horizontal pipe with main pipe
                            try
                            {
                                var c1 = Common.GetConnectorClosestTo(temp_processPipe_1, finalIntPnt);
                                var c3 = Common.GetConnectorClosestTo(horizontal_pipe, finalIntPnt);
                                if (App.m_flexSprinklerForm.IsCheckedTee)
                                {
                                    if (GetPreferredJunctionType(temp_processPipe_1) != PreferredJunctionType.Tee && isSplit == true)
                                    {
                                        CreateTap(temp_processPipe_1 as MEPCurve, horizontal_pipe as MEPCurve);
                                    }
                                    else
                                    {
                                        if (temp_processPipe_2 != null)
                                        {
                                            var c2 = Common.GetConnectorClosestTo(temp_processPipe_2, finalIntPnt);
                                            var fitting = Global.UIDoc.Document.Create.NewTeeFitting(c1, c2, c3);
                                        }
                                        else
                                        {
                                            reTrans.RollBack();
                                            continue;
                                        }
                                    }
                                }
                                else
                                {
                                    if (!isDauOng)
                                    {
                                        if (GetPreferredJunctionType(temp_processPipe_1) != PreferredJunctionType.Tee && isSplit == true)
                                        {
                                            CreateTap(temp_processPipe_1 as MEPCurve, horizontal_pipe as MEPCurve);
                                        }
                                        else
                                        {
                                            if (temp_processPipe_2 != null)
                                            {
                                                var c2 = Common.GetConnectorClosestTo(temp_processPipe_2, finalIntPnt);
                                                var fitting = Global.UIDoc.Document.Create.NewTeeFitting(c1, c2, c3);
                                            }
                                            else
                                            {
                                                reTrans.RollBack();
                                                continue;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        Global.UIDoc.Document.Create.NewElbowFitting(c1, c3);
                                    }
                                }
                            }
                            catch (System.Exception ex)
                            {
                            }

                            //  Generate vertical pipe 2
                            var line_v2 = Line.CreateBound(hor_line.GetEndPoint(1), locSprinkler);

                            var center = line_v2.Evaluate((line_v2.GetEndParameter(0) + line_v2.GetEndParameter(1)) / 2, false);

                            //Connect horizontal pipe with vertical pipe 2
                            try
                            {
                                var c1 = Common.GetConnectorClosestTo(horizontal_pipe, hor_line.GetEndPoint(1));
                                var c2 = Common.GetConnectorClosestTo(sprinkler, center);

                                FlexPipeType flexPipeType = new FilteredElementCollector(Global.UIDoc.Document).OfCategory(BuiltInCategory.OST_FlexPipeCurves).OfClass(typeof(FlexPipeType)).WhereElementIsElementType().Cast<FlexPipeType>().FirstOrDefault();
                                if (flexPipeType != null)
                                {
                                    List<XYZ> pnts = new List<XYZ>();

                                    pnts.Add(c1.Origin);
                                    pnts.Add(c2.Origin);

                                    var flexPipe = Global.UIDoc.Document.Create.NewFlexPipe(pnts, flexPipeType);

                                    //Set the diameter of flex.
                                    flexPipe.LookupParameter("Diameter").Set(App.m_flexSprinklerForm.PipeSize * Common.mmToFT);
                                    flexPipe.StartTangent = c1.CoordinateSystem.BasisZ; //Set Tangent to Conn Direction.
                                    flexPipe.EndTangent = c2.CoordinateSystem.BasisZ.Negate();

                                    var c1_flexPipe = GetConnectorClosestTo(flexPipe, c1.Origin);
                                    var c2_flexPipe = GetConnectorClosestTo(flexPipe, locSprinkler);
                                    var union = Global.UIDoc.Document.Create.NewUnionFitting(c1_flexPipe, c1);
                                    var con = Common.ToList(union.MEPModel.ConnectorManager.Connectors);

                                    var vectorMove = con[1].Origin - con[0].Origin;

                                    ElementTransformUtils.MoveElement(Global.UIDoc.Document, union.Id, vectorMove.Normalize() * 0.00001);

                                    var paraTransition = flexPipeType.get_Parameter(BuiltInParameter.RBS_CURVETYPE_DEFAULT_TRANSITION_PARAM);
                                    if (paraTransition != null && !paraTransition.IsReadOnly)
                                    {
                                        var fmlReducerSymbol = new FilteredElementCollector(Global.UIDoc.Document).OfClass(typeof(FamilySymbol))
                                                                                                                  .Cast<FamilySymbol>().FirstOrDefault(x => x.FamilyName.Contains("Reducer - Threaded - GSP"));
                                        if (fmlReducerSymbol != null)
                                            paraTransition.Set(fmlReducerSymbol.Id);
                                        Global.UIDoc.Document.Create.NewTransitionFitting(c2_flexPipe, c2);
                                    }

                                    Global.UIDoc.Document.Regenerate();
                                }
                            }
                            catch (System.Exception ex)
                            {
                            }

                            // If click cancel button when exporting
                            if (progressBar.IsCancel)
                            {
                                isCancelExport = true;
                                break;
                            }

                            nCount++;
                            dPercent = (nCount / (selSprinklers.Count * 1.0)) * 100.0;
                            progressBar.tbxMessage.Text = "Complete : " + dPercent.ToString("0.00") + "% ";
                            progressBar.IncrementProgressBar();
                        }

                        if (isCancelExport == false)
                            progressBar.Dispose();
                    }
                    catch (Exception)
                    { }
                    finally
                    {
                        progressBar.Dispose();
                    }

                    reTrans.Commit();
                }
                catch
                {
                    reTrans.RollBack();
                }
            }
            catch (Exception)
            { }
            return false;
        }

        /// <summary>
        /// Get Preferred Junction Type
        /// </summary>
        /// <param name="pipe"></param>
        /// <returns></returns>
        private static PreferredJunctionType GetPreferredJunctionType(Pipe pipe)
        {
            var pipeType = pipe.PipeType as PipeType;

            return pipeType.RoutingPreferenceManager.PreferredJunctionType;
        }

        /// <summary>
        /// Create Tap
        /// </summary>
        /// <param name="mepCurveSplit1"></param>
        /// <param name="mepCurveSplit2"></param>
        /// <returns></returns>
        public static bool CreateTap(MEPCurve mepCurveSplit1, MEPCurve mepCurveSplit2)
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
                    var tap = Global.UIDoc.Document.Create.NewTakeoffFitting(con, mepCurveSplit1);
                }
                else
                {
                    var con = GetConnectorClosestTo(mepCurveSplit2, p11);
                    var tap = Global.UIDoc.Document.Create.NewTakeoffFitting(con, mepCurveSplit1);
                }

                return true;
            }
            catch (System.Exception ex)
            {
                return false;
            }
        }

        /// <summary>
        /// Select Pipes
        /// </summary>
        /// <returns></returns>
        public static List<Pipe> SelectPipes()
        {
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

        /// <summary>
        /// Handler Process Pipes
        /// </summary>
        /// <param name="pippes"></param>
        /// <param name="sprinkle_point"></param>
        /// <param name="bSplit"></param>
        /// <returns></returns>
        private static Pipe ProcessPipes(List<Element> pipes, XYZ sprinkle_point, out bool bSplit)
        {
            bSplit = false;

            if (pipes.Count == 0)
                return null;

            Pipe pipeNear = null;

            Dictionary<Pipe, double> keyValuePairs = new Dictionary<Pipe, double>();

            foreach (Element pipe in pipes)
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
                    bSplit = true;

                var disFml = Common.ToPoint2D(p).DistanceTo(Common.ToPoint2D(sprinkle_point));

                keyValuePairs.Add(pipe as Pipe, disFml);
            }

            var min = keyValuePairs.Min(x => x.Value);

            var pairs = keyValuePairs.FirstOrDefault(x => x.Value == min);
            if (pairs.Key != null)
                pipeNear = pairs.Key;
            return pipeNear;
        }

        /// <summary>
        /// Process Start Side Pipe
        /// </summary>
        /// <param name="pipe"></param>
        /// <param name="pipe2"></param>
        /// <param name="pOn"></param>
        /// <param name="flagSplit"></param>
        public static void ProcessStartSidePipe(Pipe pipe, out Pipe pipe2, XYZ pOn, out bool isDauOng, bool flagSplit = true)
        {
            var curve = (pipe.Location as LocationCurve).Curve;

            //Create plane
            var p0 = curve.GetEndPoint(0);
            var p1 = curve.GetEndPoint(1);

            //Check co phai dau cuu hoa o gan dau cua ong ko : check trong pham vi 1m - 400mm
            double kc_mm = 400 /*1000*/;
            double km_ft = Common.mmToFT * kc_mm;

            var p02d = new XYZ(p0.X, p0.Y, 0);
            var p12d = new XYZ(p1.X, p1.Y, 0);
            var pOn2d = new XYZ(pOn.X, pOn.Y, 0);

            var d1 = p02d.DistanceTo(pOn2d);
            var d2 = p12d.DistanceTo(pOn2d);

            isDauOng = false;
            int far = -1;
            if (d1 < km_ft)
            {
                if (IsIntersect(p0) == false && !CheckPipeIsEnd(pipe, pOn) && !App.m_flexSprinklerForm.IsCheckedTee)
                {
                    isDauOng = true;
                    far = 1;
                }
            }
            else if (d2 < km_ft)
            {
                if (IsIntersect(p1) == false && !CheckPipeIsEnd(pipe, pOn) && !App.m_flexSprinklerForm.IsCheckedTee)
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
        }

        public static bool CheckPipeIsEnd(Pipe pipe, XYZ point)
        {
            var con = Common.GetConnectorClosestTo(pipe, point);

            return con.IsConnected;
        }

        /// <summary>
        /// Split Pipe
        /// </summary>
        /// <param name="pipeOrigin"></param>
        /// <param name="splitPoint"></param>
        /// <param name="pipe1"></param>
        /// <param name="pipe2"></param>
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

        /// <summary>
        /// Is Intersect
        /// </summary>
        /// <param name="point"></param>
        /// <returns></returns>
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

        /// <summary>
        /// Get Connector Closest To
        /// </summary>
        /// <param name="e"></param>
        /// <param name="p"></param>
        /// <returns></returns>
        private static Connector GetConnectorClosestTo(Element e,
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

        /// <summary>
        /// GetConnectorManager
        /// </summary>
        /// <param name="e"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentException"></exception>
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

        /// <summary>
        /// Set value to parameter based on its storage type
        /// </summary>
        private static bool SetParameterValue(Parameter param, object value)
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
        public static dynamic GetBuiltInParameterValue(Element elem, BuiltInParameter paramId)
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

        /// <summary>
        /// Get parameter value by name
        /// </summary>
        private static dynamic GetParameterValueByName(Element elem, string paramName)
        {
            if (elem != null)
            {
                Parameter parameter = elem.LookupParameter(paramName);
                return GetParameterValue(parameter);
            }
            return null;
        }
    }
}