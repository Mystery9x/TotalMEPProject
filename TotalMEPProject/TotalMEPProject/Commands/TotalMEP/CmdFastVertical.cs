﻿using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using TotalMEPProject.Ultis;

namespace TotalMEPProject.Commands.TotalMEP
{
    [Transaction(TransactionMode.Manual)]
    public class CmdFastVertical : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //if (App.CROK() != true)
            //    return Result.Failed;

            Global.UIApp = commandData.Application;
            Global.RVTApp = commandData.Application.Application;
            Global.UIDoc = commandData.Application.ActiveUIDocument;
            Global.AppCreation = commandData.Application.Application.Create;

            //Get all levels
            var coll = new FilteredElementCollector(Global.UIDoc.Document).OfClass(typeof(Level));

            if (coll.GetElementCount() == 0)
                return Result.Cancelled;

            Dictionary<string, ElementId> _levels = new Dictionary<string, ElementId>();
            foreach (Level level in coll.ToElements())
            {
                _levels.Add(level.Name, level.Id);
            }

            //Show form
            if (App.ShowFastVerticalForm(_levels) == false)
            {
                return Result.Cancelled;
            }

            return Result.Succeeded;
        }

        public static Result Process()
        {
            var fastVerticalFrm = App.fastVerticalForm;

            if (fastVerticalFrm != null && Common.IsFormSameOpen(fastVerticalFrm.Name))
                fastVerticalFrm.Hide();

            //Select pipes
            List<MEPCurve> mepCurves = SelectMEPCurves(fastVerticalFrm.GetMEPType);

            if (mepCurves == null || mepCurves.Count == 0)
            {
                fastVerticalFrm.TopMost = true;
                fastVerticalFrm.Show();
                return Result.Cancelled;
            }

            Transaction t = new Transaction(Global.UIDoc.Document, "a");

            try
            {
                t.Start();
                //Filter all slope pipes
                bool isCannot = true;
                bool isWarning = true;

                foreach (MEPCurve mepCurve in mepCurves)
                {
                    var con = Common.ToList(mepCurve.ConnectorManager.Connectors);
                    if (con?.Count > 0 && con.All(x => x.IsConnected))
                        continue;

                    var index = GetAt(mepCurve);
                    if (index == -1)
                        continue;

                    var curve = mepCurve.GetCurve();

                    var p = curve.GetEndPoint(index);
                    var pOther = curve.GetEndPoint(index == 0 ? 1 : 0);

                    var mepLevelPara = mepCurve.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM);
                    if (mepLevelPara == null)
                        continue;

                    var levelMEP = Global.UIDoc.Document.GetElement(mepLevelPara.AsElementId()) as Level;
                    if (levelMEP == null)
                        continue;

                    var offset = fastVerticalFrm.OffSet * Common.mmToFT;

                    //////////////////////////////////////////////////////////////////////////
                    bool isUp = true;

                    //////////////////////////////////////////////////////////////////////////

                    double maxZ = 0;
                    if (fastVerticalFrm.ByLevel)
                    {
                        var level = Global.UIDoc.Document.GetElement(fastVerticalFrm.LevelId) as Level;
                        if (level == null)
                            continue;

                        if (levelMEP.Elevation < level.Elevation)
                        {
                            isUp = true;
                            maxZ = level.Elevation;
                        }
                        else if (levelMEP.Elevation == level.Elevation)
                        {
                            if (p.Z > levelMEP.Elevation + offset)
                            {
                                isUp = false;
                            }
                            else
                            {
                                isUp = true;
                            }
                            maxZ = levelMEP.Elevation;
                        }
                        else
                        {
                            isUp = false;
                            maxZ = level.Elevation;
                        }

                        maxZ += offset;
                    }
                    else
                    {
                        continue;
                    }

                    //                 if (maxZ == 0)
                    //                     continue;

                    var newPlace = new XYZ(0, 0, 0);

                    var v = (p - pOther).Normalize();

                    if (fastVerticalFrm.Elbow90)
                    {
                        var level = Global.UIDoc.Document.GetElement(fastVerticalFrm.LevelId) as Level;
                        if (level == null)
                            continue;

                        var offsetFromLevel = UnitUtils.ConvertToInternalUnits(fastVerticalFrm.OffSet, DisplayUnitType.DUT_MILLIMETERS);

                        if (mepCurve is Pipe)
                        {
                            if (!Common.IsFamilySymbolSettedForPipeType(Global.UIDoc.Document, mepCurve, RoutingPreferenceRuleGroupType.Elbows))
                            {
                                //IO.ShowWarning("Chưa set RoutingPreferenceRuleGroupType !", "Warning");
                                continue;
                            }
                            Connector connector = Common.GetConnectorValid(mepCurve, level.Elevation + offsetFromLevel);

                            if (connector != null && !connector.IsConnected)
                            {
                                XYZ pointOffsetToLevel = Common.GetPointOffsetFromLevel(level, connector.Origin, fastVerticalFrm.OffSet);

                                if (pointOffsetToLevel.IsAlmostEqualTo(connector.Origin))
                                    continue;

                                if (connector.Origin.IsAlmostEqualTo(pointOffsetToLevel))
                                    continue;
                                // Tạo ống đứng vuông góc với ống ngang
                                Pipe pipeVertical = Pipe.Create(Global.UIDoc.Document, mepCurve.GetTypeId(), mepCurve.ReferenceLevel.Id, connector, pointOffsetToLevel);
                                if (pipeVertical == null)
                                    continue;

                                // Tạo kết nối giữa ống đứng và ống ngang
                                Common.GetConnectorClosedTo(mepCurve.ConnectorManager, pipeVertical.ConnectorManager, out Connector con1, out Connector con2);
                                if (con1 != null && con2 != null && Common.IsAngleValid(mepCurve, pipeVertical))
                                    Global.UIDoc.Document.Create.NewElbowFitting(con1, con2);
                                else
                                {
                                    Global.UIDoc.Document.Delete(pipeVertical.Id);

                                    if (fastVerticalFrm != null && Common.IsFormSameOpen(fastVerticalFrm.Name))
                                        fastVerticalFrm.TopMost = false;
                                    isCannot = false;
                                    continue;
                                }
                            }
                            else if (connector != null && connector.IsConnected && isWarning)
                            {
                                MessageBox.Show("Can not creat this connection for this case", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                isWarning = false;
                            }
                        }
                        else
                        {
                            var end = new XYZ(p.X, p.Y, maxZ);
                            XYZ end1 = new XYZ();

                            var elemIds = ElementTransformUtils.CopyElement(
                                Global.UIDoc.Document, mepCurve.Id, newPlace);

                            var vertical = Global.UIDoc.Document.GetElement(elemIds.ToList()[0]) as MEPCurve;

                            (vertical.Location as LocationCurve).Curve = Line.CreateBound(p, end);

                            bool isCop = false;
                            Connector c = null;
                            foreach (Connector con1 in mepCurve.ConnectorManager.Connectors)
                            {
                                if (con1.IsConnected)
                                {
                                    isCop = true;
                                    continue;
                                }
                                else
                                    c = con1;
                            }

                            if (isCop)
                            {
                                end1 = new XYZ(c.Origin.X, c.Origin.Y, maxZ);
                                (vertical.Location as LocationCurve).Curve = Line.CreateBound(c.Origin, end1);
                            }

                            if (vertical as CableTray != null)
                            {
                                SetCurveNormal(vertical as CableTray, mepCurve as CableTray);
                            }
                            else if (vertical as Autodesk.Revit.DB.Mechanical.Duct != null)
                            {
                                if (vertical as Autodesk.Revit.DB.Mechanical.Duct != null)
                                {
                                    //Rotate
                                    var a = v.AngleOnPlaneTo(XYZ.BasisY, XYZ.BasisZ);
                                    if (!isCop)
                                    {
                                        var lv = Line.CreateBound(p, end);

                                        ElementTransformUtils.RotateElement(Global.UIDoc.Document, vertical.Id, lv, a);
                                    }
                                    else
                                    {
                                        var lv = Line.CreateBound(c.Origin, end1);

                                        ElementTransformUtils.RotateElement(Global.UIDoc.Document, vertical.Id, lv, a);
                                    }
                                }
                            }

                            var elbow = CreateElbow(mepCurve, vertical);
                        }
                    }
                    else if (fastVerticalFrm.Elbow45)
                    {
                        if (!Common.IsFamilySymbolSettedForPipeType(Global.UIDoc.Document, mepCurve, RoutingPreferenceRuleGroupType.Elbows))
                            continue;

                        var pipeType = Global.UIDoc.Document.GetElement(mepCurve.GetTypeId()) as PipeType;
                        if (pipeType == null)
                            continue;
                        var level = Global.UIDoc.Document.GetElement(fastVerticalFrm.LevelId) as Level;
                        if (level == null)
                            continue;

                        double slope = Math.Round((double)Common.GetValueParameterByBuilt(mepCurve, BuiltInParameter.RBS_PIPE_SLOPE), 5);
                        Pipe pipeVertical = null;
                        if (pipeType.Name.Contains(Define.PIPE_CAST_IRON))
                        {
                            if (Common.IsEqual(slope, 0))
                            {
                                Connector conHor1 = Common.GetConnectorValid1(mepCurve);
                                if (conHor1 != null && !conHor1.IsConnected)
                                {
                                    conHor1 = Common.GetConnectorNearest(conHor1.Origin, mepCurve, out Connector conHor2);

                                    XYZ pointOffsetToLevel = Common.GetPointOffsetFromLevel(level, conHor1.Origin, fastVerticalFrm.OffSet);

                                    if (pointOffsetToLevel.IsAlmostEqualTo(conHor1.Origin))
                                        continue;

                                    pipeVertical = Pipe.Create(Global.UIDoc.Document, mepCurve.GetTypeId(), mepCurve.ReferenceLevel.Id, conHor1, pointOffsetToLevel);
                                    if (pipeVertical == null)
                                        continue;
                                }
                            }
                            else
                            {
                                XYZ pointSt = ((LocationCurve)mepCurve.Location).Curve.GetEndPoint(0);

                                Connector conSt = Common.GetConnectorNearest(pointSt, mepCurve, out Connector conEnd);

                                XYZ pointOffsetToLevelSt = Common.GetPointOffsetFromLevel(level, conSt.Origin, fastVerticalFrm.OffSet);

                                XYZ pointOffsetToLevelEnd = Common.GetPointOffsetFromLevel(level, conEnd.Origin, fastVerticalFrm.OffSet);

                                bool isTop = false;
                                if (pointOffsetToLevelSt.Z > conSt.Origin.Z && pointOffsetToLevelEnd.Z > conEnd.Origin.Z)
                                    isTop = true;

                                Connector retval = null;
                                if (!conSt.IsConnected && !conEnd.IsConnected)
                                {
                                    if (isTop)
                                        retval = (conSt.Origin.Z > conEnd.Origin.Z) ? conSt : conEnd;
                                    else
                                        retval = (conSt.Origin.Z < conEnd.Origin.Z) ? conSt : conEnd;

                                    pointOffsetToLevelSt = Common.GetPointOffsetFromLevel(level, retval.Origin, fastVerticalFrm.OffSet);
                                    if (pointOffsetToLevelSt.IsAlmostEqualTo(retval.Origin))
                                        continue;

                                    pipeVertical = Pipe.Create(Global.UIDoc.Document, mepCurve.GetTypeId(), mepCurve.ReferenceLevel.Id, retval, pointOffsetToLevelSt);
                                    if (pipeVertical == null)
                                        continue;
                                }
                                else
                                {
                                    Connector conHor1 = Common.GetConnectorValid1(mepCurve);
                                    if (conHor1 != null && !conHor1.IsConnected)
                                    {
                                        conHor1 = Common.GetConnectorNearest(conHor1.Origin, mepCurve, out Connector conHor2);

                                        XYZ pointOffsetToLevel = Common.GetPointOffsetFromLevel(level, conHor1.Origin, fastVerticalFrm.OffSet);

                                        if (pointOffsetToLevel.IsAlmostEqualTo(conHor1.Origin))
                                            continue;

                                        pipeVertical = Pipe.Create(Global.UIDoc.Document, mepCurve.GetTypeId(), mepCurve.ReferenceLevel.Id, conHor1, pointOffsetToLevel);
                                        if (pipeVertical == null)
                                            continue;
                                    }
                                }
                            }

                            Common.GetConnectorClosedTo(mepCurve.ConnectorManager, pipeVertical.ConnectorManager, out Connector con1, out Connector con2);
                            if (con1 != null && con2 != null && con1.IsConnectedTo(con2))
                                con1.DisconnectFrom(con2);
                            Common.ConnectPipeVerticalElbow45(Global.UIDoc.Document, mepCurve, pipeVertical, true);
                        }
                        else
                        {
                            var offsetFromLevel = UnitUtils.ConvertToInternalUnits(fastVerticalFrm.OffSet, DisplayUnitType.DUT_MILLIMETERS);
                            //Tao mot pipe 45
                            var elemIds = ElementTransformUtils.CopyElement(
                                Global.UIDoc.Document, mepCurve.Id, newPlace);

                            var vertical45_2 = Global.UIDoc.Document.GetElement(elemIds.ToList()[0]) as MEPCurve;

                            Connector conHor1 = Common.GetConnectorValid(mepCurve, level.Elevation + offsetFromLevel);
                            if (conHor1 != null && !conHor1.IsConnected)
                            {
                                conHor1 = Common.GetConnectorNearest(conHor1.Origin, mepCurve, out Connector conHor2);

                                XYZ pointOffsetToLevel = Common.GetPointOffsetFromLevel(level, conHor1.Origin, fastVerticalFrm.OffSet);

                                if (pointOffsetToLevel.IsAlmostEqualTo(conHor1.Origin))
                                    continue;

                                Line line = Line.CreateBound(conHor1.Origin, pointOffsetToLevel);
                                (vertical45_2.Location as LocationCurve).Curve = line;

                                Common.ConnectPipeVerticalElbow45(Global.UIDoc.Document, mepCurve, vertical45_2, true);
                            }
                            else
                                Global.UIDoc.Document.Delete(vertical45_2.Id);
                        }
                    }
                    else if (fastVerticalFrm.Siphon)
                    {
                        if (mepCurve is Pipe)
                        {
                            var offsetFromLevel = UnitUtils.ConvertToInternalUnits(fastVerticalFrm.OffSet, DisplayUnitType.DUT_MILLIMETERS);
                            var symbol = Global.UIDoc.Document.GetElement(fastVerticalFrm.SiphonId) as FamilySymbol;

                            if (symbol.IsActive == false)
                                symbol.Activate();

                            if (Common.IsPipeVertical(mepCurve))
                                continue;
                            var level = Global.UIDoc.Document.GetElement(fastVerticalFrm.LevelId) as Level;
                            if (level == null)
                                continue;
                            Connector connector = Common.GetConnectorValid(mepCurve, level.Elevation + offsetFromLevel);
                            if (connector != null && !connector.IsConnected)
                            {
                                XYZ pointOffsetToLevel = Common.GetPointOffsetFromLevel(level, connector.Origin, fastVerticalFrm.OffSet);

                                if (pointOffsetToLevel.IsAlmostEqualTo(connector.Origin))
                                    continue;

                                //XYZ vectorMoveZ = (pointOffsetToLevel - connector.Origin).Normalize();

                                //if (vectorMoveZ.IsAlmostEqualTo(XYZ.BasisZ.Negate()))
                                //    continue;

                                if (connector.Origin.IsAlmostEqualTo(pointOffsetToLevel))
                                    continue;

                                // Create pipe vertical , it perpendicular with pipe horizontal
                                Pipe pipeVertical = Pipe.Create(Global.UIDoc.Document, mepCurve.GetTypeId(), mepCurve.ReferenceLevel.Id, connector, pointOffsetToLevel);

                                if (pipeVertical != null)
                                {
                                    Common.GetConnectorClosedTo(pipeVertical.ConnectorManager, mepCurve.ConnectorManager, out Connector conSt, out Connector conEnd);

                                    if (conSt.IsConnectedTo(conEnd))
                                        conSt.DisconnectFrom(conEnd);

                                    CreateSiphon(mepCurve, pipeVertical, fastVerticalFrm.SiphonId);
                                    //ConnectPipeVerticalSiphon(Global.UIDoc.Document, pipeVertical, mepCurve, symbol);
                                }
                            }
                        }
                        else
                        {
                            bool isCop = false;
                            Connector c = null;
                            XYZ end1 = new XYZ();
                            var end = new XYZ(p.X, p.Y, maxZ);

                            var elemIds = ElementTransformUtils.CopyElement(
                                Global.UIDoc.Document, mepCurve.Id, newPlace);

                            // Create a Outline, uses a minimum and maximum XYZ point to initialize the outline.
                            Outline myOutLn = CreateOutLineFromBoundingBox(mepCurve);
                            if (myOutLn == null || myOutLn.IsEmpty)
                                continue;

                            // Create a BoundingBoxIntersects filter with this Outline
                            BoundingBoxIntersectsFilter filter = new BoundingBoxIntersectsFilter(myOutLn);

                            FilteredElementCollector collector = new FilteredElementCollector(Global.UIDoc.Document);

                            var fittingConnected = collector.WherePasses(filter).Where(x => x is FamilyInstance).ToList();

                            if (fittingConnected?.Count == 1)
                            {
                                var loc = fittingConnected.FirstOrDefault().Location as LocationPoint;
                                if (loc == null)
                                    continue;

                                foreach (Connector con1 in mepCurve.ConnectorManager.Connectors)
                                {
                                    if (con1.Origin.IsAlmostEqualTo(loc.Point))
                                    {
                                        isCop = true;
                                        continue;
                                    }
                                    else
                                        c = con1;
                                }
                            }
                            else if (fittingConnected?.Count > 1)
                                continue;

                            var vertical = Global.UIDoc.Document.GetElement(elemIds.ToList()[0]) as MEPCurve;

                            (vertical.Location as LocationCurve).Curve = Line.CreateBound(p, end);

                            if (isCop)
                            {
                                end1 = new XYZ(c.Origin.X, c.Origin.Y, maxZ);
                                (vertical.Location as LocationCurve).Curve = Line.CreateBound(c.Origin, end1);
                            }

                            if (vertical as CableTray != null)
                            {
                                SetCurveNormal(vertical as CableTray, mepCurve as CableTray);
                            }
                            else if (vertical as Autodesk.Revit.DB.Mechanical.Duct != null)
                            {
                                //Can rotate vertical duct
                                if (vertical as Autodesk.Revit.DB.Mechanical.Duct != null)
                                {
                                    //Rotate
                                    var a = v.AngleOnPlaneTo(XYZ.BasisY, XYZ.BasisZ);
                                    if (!isCop)
                                    {
                                        var lv = Line.CreateBound(p, end);

                                        ElementTransformUtils.RotateElement(Global.UIDoc.Document, vertical.Id, lv, a);
                                    }
                                    else
                                    {
                                        var lv = Line.CreateBound(c.Origin, end1);

                                        ElementTransformUtils.RotateElement(Global.UIDoc.Document, vertical.Id, lv, a);
                                    }
                                }
                            }

                            CreateSiphon(mepCurve, vertical, fastVerticalFrm.SiphonId);
                        }
                    }
                }
                if (!isCannot)
                    IO.ShowWarning(Define.ERR_CAN_NOT_CREATE_THIS_CONNECTION_FOR_THIS_CASE, "Warning");
                t.Commit();
            }
            catch (Exception ex)
            {
                string mess = ex.Message;

                if (fastVerticalFrm != null && Common.IsFormSameOpen(fastVerticalFrm.Name))
                    fastVerticalFrm.TopMost = false;
            }

            if (fastVerticalFrm != null && Common.IsFormSameOpen(fastVerticalFrm.Name))
            {
                fastVerticalFrm.TopMost = true;
                fastVerticalFrm.Show();
                Global.UIDoc.RefreshActiveView();
            }

            if (t.HasStarted())
            {
                t.RollBack();
            }

            return Result.Succeeded;
        }

        /// <summary>
        /// CreateOutLineFromBoundingBox
        /// </summary>
        /// <param name="ele"></param>
        /// <returns></returns>

        public static Outline CreateOutLineFromBoundingBox(Element ele)
        {
            Outline retVal = null;
            if (ele == null)
                return retVal;
            BoundingBoxXYZ boundingBox = ele.get_BoundingBox(null);
            if (boundingBox == null)
                return retVal;
            XYZ min = new XYZ(boundingBox.Min.X, boundingBox.Min.Y, boundingBox.Min.Z);
            XYZ max = new XYZ(boundingBox.Max.X, boundingBox.Max.Y, boundingBox.Max.Z);
            retVal = new Outline(min, max);
            return retVal;
        }

        private static FamilyInstance CreateSiphon(MEPCurve mepCurve, MEPCurve vertical_pipe, ElementId symbolId)
        {
            try
            {
                var symbol = Global.UIDoc.Document.GetElement(symbolId) as FamilySymbol;

                if (symbol.IsActive == false)
                    symbol.Activate();

                var curve = (vertical_pipe.Location as LocationCurve).Curve;
                var p0 = curve.GetEndPoint(0);

                var mepLevelPara = vertical_pipe.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM);
                if (mepLevelPara == null)
                    return null;

                var levelMEP = Global.UIDoc.Document.GetElement(mepLevelPara.AsElementId()) as Level;
                if (levelMEP == null)
                    return null;

                SubTransaction sub = new SubTransaction(Global.UIDoc.Document);
                sub.Start();

                FamilyInstance fitting = null;
                try
                {
                    fitting = Global.UIDoc.Document.Create.NewFamilyInstance(p0, symbol, levelMEP, StructuralType.NonStructural);

                    sub.Commit();
                }
                catch (System.Exception ex)
                {
                    sub.RollBack();
                    return null;
                }
                Global.UIDoc.Document.Regenerate();

                //////////////////////////////////////////////////////////////////////////
                //try
                //{
                //    var axis = Line.CreateUnbound(p0, XYZ.BasisZ);
                //    (fiting.Location).Rotate(axis, -Math.PI / 2);
                //}
                //catch (System.Exception ex)
                //{
                //}

                //try
                //{
                //    var axis = Line.CreateUnbound(p0, XYZ.BasisX);
                //    (fiting.Location).Rotate(axis, Math.PI / 2);
                //}
                //catch (System.Exception ex)
                //{
                //}
                //////////////////////////////////////////////////////////////////////////

                sub.Start();
                var system = fitting.MEPModel as MechanicalFitting;

                var list = Common.ToList(system.ConnectorManager.Connectors);
                if (list.Count == 2)
                {
                    try
                    {
                        var axis = Line.CreateUnbound(p0, XYZ.BasisZ);
                        (fitting.Location).Rotate(axis, -Math.PI / 2);
                    }
                    catch (System.Exception ex)
                    {
                    }
                    (fitting.Location as LocationPoint).Point = p0;
                    try
                    {
                        var axis = Line.CreateUnbound(p0, XYZ.BasisX);
                        (fitting.Location).Rotate(axis, Math.PI / 2);
                    }
                    catch (System.Exception ex)
                    {
                    }
                     (fitting.Location as LocationPoint).Point = p0;

                    var cPipe = Common.GetConnectorClosestTo(mepCurve, p0);
                    var con = Common.GetConnectorClosestTo(fitting, p0);

                    //further
                    con = list[0].Id == con.Id ? list[1] : list[0];
                    //////////////////////////////////////////////////////////////////////////
                    (fitting.Location as LocationPoint).Point = p0;

                    var v_to_main = con.CoordinateSystem.BasisZ;
                    var v_pipe = cPipe.CoordinateSystem.BasisZ;

                    var angle = XYZ.BasisX.AngleOnPlaneTo(v_pipe, XYZ.BasisZ);// v_pipe.AngleOnPlaneTo(v_to_main, XYZ.BasisZ);
                    try
                    {
                        var axis = Line.CreateUnbound(p0, XYZ.BasisZ);
                        (fitting.Location).Rotate(axis, angle + Math.PI);
                    }
                    catch (System.Exception ex)
                    {
                    }

                    Global.UIDoc.Document.Regenerate();
                    (fitting.Location as LocationPoint).Point = p0;

                    //////////////////////////////////////////////////////////////////////////
                    //////////////////////////////////////////////////////////////////////////

                    con.ConnectTo(cPipe);
                    if (cPipe.IsConnectedTo(con))
                    {
                        try
                        {
                            con.Radius = cPipe.Radius;
                        }
                        catch (System.Exception ex)
                        {
                        }
                    }
                    Global.UIDoc.Document.Regenerate();
                    (fitting.Location as LocationPoint).Point = p0;

                    var cPipe_vertical = Common.GetConnectorClosestTo(vertical_pipe, p0);

                    var other = list[0].Id == con.Id ? list[1] : list[0];

                    other.ConnectTo(cPipe_vertical);
                    if (cPipe_vertical.IsConnectedTo(other))
                    {
                        try
                        {
                            other.Radius = cPipe_vertical.Radius;
                        }
                        catch (System.Exception ex)
                        {
                        }
                    }
                }

                //NOTE: Phai set 1 giá trị location de no tu dong quay theo chieu cua pipe : chú ý, vì đây là giải pháp tạm thời
                (fitting.Location as LocationPoint).Point = new XYZ(p0.X, p0.Y, p0.Z + 0.001);
                Global.UIDoc.Document.Regenerate();
                (fitting.Location as LocationPoint).Point = p0;
                //END : NOTE////////////////////////////////////////////////////////////////////////

                sub.Commit();
                return fitting;
            }
            catch (System.Exception ex)
            {
                string mess = ex.Message;
                return null;
            }
        }

        /// <summary>
        /// Kết nối ống nằm ngang với ống đứng bằng kết nối dạng siphon
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="pipe1"></param>
        /// <param name="pipe2"></param>
        /// <param name="typeSiphon"></param>
        public static void ConnectPipeVerticalSiphon(Document doc, MEPCurve pipe1, MEPCurve pipe2, FamilySymbol typeSiphon)
        {
            if (doc == null || pipe1 == null || pipe2 == null)
                return;
            try
            {
                if (!typeSiphon.IsActive)
                    typeSiphon.Activate();

                if (pipe1 != null && pipe2 != null && pipe1.Location is LocationCurve && pipe2.Location is LocationCurve)
                {
                    bool isSucces = Common.DetermindPipeVertcalAndHorizontal(pipe1, pipe2, out MEPCurve pipeVertical, out MEPCurve pipeHorizontal);

                    Common.GetConnectorClosedTo(pipeVertical.ConnectorManager, pipeHorizontal.ConnectorManager, out Connector conSt, out Connector conEnd);

                    if (conSt == null || conSt.IsConnected || conEnd == null || conEnd.IsConnected)
                        return;

                    if (isSucces)
                    {
                        // Di chuyển 2 pipe về cùng tâm
                        Common.MovePipeToCenter(doc, pipeVertical, pipeHorizontal);

                        Connector conLower = Common.GetConnectorMinZ(pipeVertical, out Connector conHigher);
                        if (conLower == null)
                            return;

                        //Ngắt kết nối với fitting
                        Connector conStPipeHor = pipeHorizontal.ConnectorManager.Lookup(0);
                        Connector conEndPipeHor = pipeHorizontal.ConnectorManager.Lookup(1);

                        FamilyInstance elbowFittingConnectedHor = null;
                        if (conStPipeHor != null && conEndPipeHor != null)
                        {
                            if (conStPipeHor.IsConnected || conEndPipeHor.IsConnected)
                            {
                                Connector conNotConnectedHor = (conStPipeHor.IsConnected) ? conStPipeHor : conEndPipeHor;

                                elbowFittingConnectedHor = Common.GetFittingConnected(doc, pipeHorizontal, conNotConnectedHor);
                                if (elbowFittingConnectedHor != null)
                                {
                                    Common.GetConnectorClosedTo(pipeHorizontal.ConnectorManager, elbowFittingConnectedHor.MEPModel.ConnectorManager, out Connector conPipeHor1, out Connector conFitHor);
                                    if (conPipeHor1 != null && conFitHor != null && conPipeHor1.IsConnectedTo(conFitHor))
                                        conPipeHor1.DisconnectFrom(conFitHor);
                                }
                            }
                        }

                        //Connector conHorCheck = ConnectorUtils.GetConnectorNearest(conHigher.Origin, pipeHorizontal, out Connector conUpper);

                        //if (conHorCheck.Origin.Z > conHigher.Origin.Z || Common.IsEqual(conHorCheck.Origin.Z, conHigher.Origin.Z))
                        //{
                        //    IO.ShowWarning(Define.ERR_CAN_NOT_CREATE_THIS_CONNECTION_FOR_THIS_CASE, "Warning");
                        //    return;
                        //}

                        if (pipeHorizontal.Diameter != pipeVertical.Diameter)
                            Common.SetValueParameterByBuiltIn(pipeHorizontal, BuiltInParameter.RBS_PIPE_DIAMETER_PARAM, pipeVertical.Diameter);

                        FamilyInstance fittingSiphon = doc.Create.NewFamilyInstance(conLower.Origin, typeSiphon, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                        if (fittingSiphon != null)
                        {
                            if (pipeHorizontal.Name.Equals(Define.PIPE_CAST_IRON))
                            {
                                //Đặt lại bán kính cho siphon
                                Common.SetValueParameterByListName(fittingSiphon, pipeVertical.Diameter / 2, "RND");

                                if (!Common.IsSiphonVertical(fittingSiphon))
                                {
                                    Line axisY = Line.CreateBound(conLower.Origin, Common.GetPointOnVector(conLower.Origin, XYZ.BasisY, 10));
                                    fittingSiphon.Location.Rotate(axisY, Math.PI / 2);
                                }

                                double angle = Common.GetAngleBetweenPipeAndSiphon(pipeHorizontal, fittingSiphon);

                                Line axisZ = Line.CreateBound(conLower.Origin, Common.GetPointOnVector(conLower.Origin, XYZ.BasisZ, 10));
                                fittingSiphon.Location.Rotate(axisZ, -angle);

                                Connector conSiphonTop = fittingSiphon.MEPModel.ConnectorManager.Lookup(1);
                                Connector conSiphonRight = fittingSiphon.MEPModel.ConnectorManager.Lookup(2);

                                if (conSiphonTop == null || conSiphonRight == null)
                                    return;

                                // Kết nối siphon với ống thẳng đứng

                                Connector con1 = Common.GetConnectorNearest(conSiphonTop.Origin, pipeVertical, out Connector con2);

                                Common.ResetLocation(pipeVertical, conSiphonTop.Origin, con2.Origin);

                                Common.GetConnectorClosedTo(pipeVertical.ConnectorManager, fittingSiphon.MEPModel.ConnectorManager, out con1, out con2);
                                con1.ConnectTo(con2);

                                // Di chuyển pipe nằm ngang đến vị trí mới

                                // Ống nằm ngang
                                Connector conHor1 = Common.GetConnectorNearest(conSiphonRight.Origin, pipeHorizontal, out Connector conHor2);

                                XYZ vectorMoveZ1 = (conHor1.Origin.Z > conHor2.Origin.Z || Common.IsEqual(conHor1.Origin.Z, conHor2.Origin.Z)) ? XYZ.BasisZ : XYZ.BasisZ.Negate();

                                XYZ newPoint2 = new XYZ(conSiphonRight.Origin.X, conSiphonRight.Origin.Y, conHor2.Origin.Z);
                                double distance = newPoint2.DistanceTo(conHor2.Origin);
                                double slope = Math.Round((double)Common.GetValueParameterByBuilt(pipeHorizontal, BuiltInParameter.RBS_PIPE_SLOPE), 5);
                                newPoint2 = Common.GetPointOnVector(newPoint2, vectorMoveZ1, slope * distance);

                                // Kéo dài ống thẳng đứng và siphon tới ống nằm ngang
                                double height = Math.Abs(conSiphonRight.Origin.Z - newPoint2.Z);

                                XYZ vectran = (XYZ.BasisZ.Negate()) * height;

                                if (newPoint2.Z > conSiphonRight.Origin.Z)
                                    ElementTransformUtils.MoveElement(doc, fittingSiphon.Id, vectran.Negate());
                                else
                                    ElementTransformUtils.MoveElement(doc, fittingSiphon.Id, vectran);

                                Common.ResetLocation(pipeHorizontal, conHor2.Origin, newPoint2);

                                // Kết nối ống nằm ngang với siphon
                                Common.GetConnectorClosedTo(pipeHorizontal.ConnectorManager, fittingSiphon.MEPModel.ConnectorManager, out con1, out con2);
                                con1.ConnectTo(con2);

                                if (Common.IsEqual(slope, 0))
                                {
                                    Connector conStPipe = pipeHorizontal.ConnectorManager.Lookup(0);
                                    Connector conEndPipe = pipeHorizontal.ConnectorManager.Lookup(1);

                                    if (conStPipe != null && conEndPipe != null)
                                    {
                                        if (!conStPipe.IsConnected || !conEndPipe.IsConnected)
                                        {
                                            Connector conPipeNotConnected = (!conStPipe.IsConnected) ? conStPipe : conEndPipe;

                                            FamilyInstance elbowFittingConnected = Common.GetFittingConnected(doc, pipeHorizontal, conPipeNotConnected, fittingSiphon.Id);
                                            if (elbowFittingConnected != null)
                                            {
                                                Common.GetConnectorClosedTo(pipeHorizontal.ConnectorManager, elbowFittingConnected.MEPModel.ConnectorManager, out con1, out con2);
                                                if (con1 != null && con2 != null && !con1.IsConnected && !con2.IsConnected)
                                                    con1.ConnectTo(con2);
                                            }
                                        }
                                    }
                                }
                            }
                            else
                            {
                                //Đặt lại bán kính cho siphon
                                Common.SetValueParameterByListName(fittingSiphon, pipeVertical.Diameter, "REDY_BDN_DN", "ENG_DN");

                                // Quay siphon về ống nằm ngang
                                Line axisY = Line.CreateBound(conLower.Origin, Common.GetPointOnVector(conLower.Origin, XYZ.BasisY, 10));
                                fittingSiphon.Location.Rotate(axisY, Math.PI / 2);

                                double angle = Common.GetAngleBetweenPipeAndSiphon(pipeHorizontal, fittingSiphon);

                                Line axisZ = Line.CreateBound(conLower.Origin, Common.GetPointOnVector(conLower.Origin, XYZ.BasisZ, 10));
                                fittingSiphon.Location.Rotate(axisZ, -angle);

                                // Kết nối siphon với ống thẳng đứng
                                Common.GetConnectorClosedTo(pipeVertical.ConnectorManager, fittingSiphon.MEPModel.ConnectorManager, out Connector con1, out Connector con2);
                                con1.ConnectTo(con2);

                                // Di chuyển pipe nằm ngang đến vị trí mới
                                Connector conNotConnected = fittingSiphon.MEPModel.ConnectorManager.Lookup(1);
                                if (conNotConnected == null)
                                    return;

                                // Ống nằm ngang
                                Connector conHor1 = Common.GetConnectorNearest(conNotConnected.Origin, pipeHorizontal, out Connector conHor2);
                                XYZ vectorMoveZ1 = (conHor1.Origin.Z > conHor2.Origin.Z || Common.IsEqual(conHor1.Origin.Z, conHor2.Origin.Z)) ? XYZ.BasisZ : XYZ.BasisZ.Negate();
                                XYZ newPoint2 = new XYZ(conNotConnected.Origin.X, conNotConnected.Origin.Y, conHor2.Origin.Z);
                                double distance = newPoint2.DistanceTo(conHor2.Origin);
                                double slope = Math.Round((double)Common.GetValueParameterByBuilt(pipeHorizontal, BuiltInParameter.RBS_PIPE_SLOPE), 5);
                                newPoint2 = Common.GetPointOnVector(newPoint2, vectorMoveZ1, slope * distance);

                                double height = Math.Abs(conNotConnected.Origin.Z - newPoint2.Z);

                                XYZ vectran = (XYZ.BasisZ.Negate()) * height;

                                if (newPoint2.Z > conNotConnected.Origin.Z)
                                    ElementTransformUtils.MoveElement(doc, fittingSiphon.Id, vectran.Negate());
                                else
                                    ElementTransformUtils.MoveElement(doc, fittingSiphon.Id, vectran);

                                Common.ResetLocation(pipeHorizontal, conHor2.Origin, newPoint2);

                                // Kết nối ống nằm ngang với siphon
                                Common.GetConnectorClosedTo(pipeHorizontal.ConnectorManager, fittingSiphon.MEPModel.ConnectorManager, out con1, out con2);
                                con1.ConnectTo(con2);

                                if (Common.IsEqual(slope, 0))
                                {
                                    Connector conStPipe = pipeHorizontal.ConnectorManager.Lookup(0);
                                    Connector conEndPipe = pipeHorizontal.ConnectorManager.Lookup(1);

                                    if (conStPipe != null && conEndPipe != null)
                                    {
                                        if (!conStPipe.IsConnected || !conEndPipe.IsConnected)
                                        {
                                            Connector conPipeNotConnected = (!conStPipe.IsConnected) ? conStPipe : conEndPipe;

                                            FamilyInstance elbowFittingConnected = Common.GetFittingConnected(doc, pipeHorizontal, conPipeNotConnected, fittingSiphon.Id);
                                            if (elbowFittingConnected != null)
                                            {
                                                Common.GetConnectorClosedTo(pipeHorizontal.ConnectorManager, elbowFittingConnected.MEPModel.ConnectorManager, out con1, out con2);
                                                if (con1 != null && con2 != null && !con1.IsConnected && !con2.IsConnected)
                                                    con1.ConnectTo(con2);
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        // Join lại fitting
                        if (elbowFittingConnectedHor != null)
                        {
                            Common.GetConnectorClosedTo(pipeHorizontal.ConnectorManager, elbowFittingConnectedHor.MEPModel.ConnectorManager, out Connector conPipeHor1, out Connector conFitHor);
                            if (conPipeHor1 != null && conFitHor != null && !conPipeHor1.IsConnected && !conFitHor.IsConnected)
                            {
                                elbowFittingConnectedHor.Location.Move(conPipeHor1.Origin - conFitHor.Origin);

                                conPipeHor1.ConnectTo(conFitHor);
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                //IO.ShowError(ex.Message, "Error");
                return;
            }
        }

        private static void SetCurveNormal(CableTray vertical, CableTray current)
        {
            var curve = current.GetCurve() as Line;
            if (curve == null)
                return;

            var v = curve.Direction;

            (vertical as CableTray).CurveNormal = v;
        }

        public static List<MEPCurve> SelectMEPCurves(Type type)
        {
            //Pick pipe
            List<MEPCurve> pipes = new List<MEPCurve>();
            try
            {
                var pickedObjs = Global.UIDoc.Selection.PickObjects(ObjectType.Element, new MEPCurveFilter/*PipeFilter*/(type), "Pick mep elements: ");

                foreach (Reference pickedObj in pickedObjs)
                {
                    var pipe = Global.UIDoc.Document.GetElement(pickedObj) as MEPCurve;

                    if (pipe != null)
                        pipes.Add(pipe);
                }
            }
            catch (System.Exception ex)
            {
            }

            return pipes;
        }

        private void MoveToOldPosition(List<ElementId> ids, Line old_curve, XYZ oldPoint, XYZ current)
        {
            var d = (Common.ToPoint2D(oldPoint) - Common.ToPoint2D(current)).GetLength();

            var newLocation = old_curve.Evaluate(d, false);

            XYZ transform = newLocation - oldPoint;
            ElementTransformUtils.MoveElements(Global.UIDoc.Document, ids, transform);
        }

        private static FamilyInstance CreateElbow(MEPCurve pipe1, MEPCurve pipe2)
        {
            try
            {
                List<Connector> connectors = Common.GetConnectionNearest(pipe1, pipe2);
                if (connectors != null && connectors.Count == 2)
                {
                    var elbow = Global.UIDoc.Document.Create.NewElbowFitting(connectors[0], connectors[1]);

                    return elbow;
                }
            }
            catch (System.Exception ex)
            {
            }
            return null;
        }

        private static int GetAt(MEPCurve pipe)
        {
            var curve = (pipe.Location as LocationCurve).Curve;

            var p0 = curve.GetEndPoint(0);
            var p1 = curve.GetEndPoint(1);

            bool atStart = false;

            atStart = CheckIntersection(p0, pipe);

            bool atEnd = false;

            atEnd = CheckIntersection(p1, pipe);

            if (atStart == true && atEnd == false)
                return 1;
            else if (atStart == false && atEnd == true)
                return 0;
            else if (atStart == false && atEnd == false) // Ko intersection thi uu tien tai end
                return 1;

            return -1;
        }

        private static bool CheckIntersection(XYZ p, MEPCurve pipe)
        {
            double ft = 0.1;
            var solid = Common.CreateCylindricalVolume(p, ft, ft, false);
            if (solid != null)
            {
                //Find intersection
                FilteredElementCollector collector = new FilteredElementCollector(Global.UIDoc.Document);
                collector.Excluding(new List<ElementId> { pipe.Id });

                if (pipe is Pipe)
                    collector.OfClass(typeof(Pipe));
                else if (pipe is Autodesk.Revit.DB.Mechanical.Duct)
                    collector.OfClass(typeof(Autodesk.Revit.DB.Mechanical.Duct));
                else if (pipe is CableTray)
                    collector.OfClass(typeof(CableTray));
                else if (pipe is Conduit)
                    collector.OfClass(typeof(Conduit));

                collector.WherePasses(new ElementIntersectsSolidFilter(solid)); // Apply intersection filter to find matches

                if (collector.GetElementCount() == 0)
                    return false;

                var intersections = collector.ToElements();

                //Find fitting
                collector = new FilteredElementCollector(Global.UIDoc.Document);
                collector.Excluding(new List<ElementId> { pipe.Id });

                if (pipe is Pipe)
                    collector.OfCategory(BuiltInCategory.OST_PipeFitting);
                else if (pipe is Autodesk.Revit.DB.Mechanical.Duct)
                    collector.OfCategory(BuiltInCategory.OST_DuctFitting);
                else if (pipe is CableTray)
                    collector.OfCategory(BuiltInCategory.OST_CableTrayFitting);
                else if (pipe is Conduit)
                    collector.OfCategory(BuiltInCategory.OST_ConduitFitting);

                collector.WherePasses(new ElementIntersectsSolidFilter(solid)); // Apply intersection filter to find matches

                if (collector.GetElementCount() == 0)
                    return false;

                return true;
            }
            return false;
        }
    }
}