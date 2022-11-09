using Autodesk.Revit.Attributes;
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
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TotalMEPProject.UI;
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

            //Select pipes
            List<MEPCurve> mepCurves = SelectMEPCurves(fastVerticalFrm.GetMEPType);

            if (mepCurves == null || mepCurves.Count == 0)
                return Result.Cancelled;

            Transaction t = new Transaction(Global.UIDoc.Document, "a");

            try
            {
                t.Start();
                //Filter all slope pipes
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
                        var end = new XYZ(p.X, p.Y, maxZ);

                        var elemIds = ElementTransformUtils.CopyElement(
                            Global.UIDoc.Document, mepCurve.Id, newPlace);

                        var vertical = Global.UIDoc.Document.GetElement(elemIds.ToList()[0]) as MEPCurve;

                        (vertical.Location as LocationCurve).Curve = Line.CreateBound(p, end);

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
                                var lv = Line.CreateBound(p, end);

                                ElementTransformUtils.RotateElement(Global.UIDoc.Document, vertical.Id, lv, a);
                            }
                        }

                        var elbow = CreateElbow(mepCurve, vertical);
                    }
                    else if (fastVerticalFrm.Elbow45)
                    {
                        var pipeType = Global.UIDoc.Document.GetElement(mepCurve.GetTypeId()) as PipeType;
                        if (pipeType == null)
                            continue;

                        if (pipeType.Name.Contains(Define.PIPE_CAST_IRON))
                        {
                            var level = Global.UIDoc.Document.GetElement(fastVerticalFrm.LevelId) as Level;
                            if (level == null)
                                continue;
                            CreateElbowPipeIron(mepCurve, level, fastVerticalFrm.OffSet);
                        }
                        else
                        {
                            //Tao mot pipe 45
                            var elemIds = ElementTransformUtils.CopyElement(
                                Global.UIDoc.Document, mepCurve.Id, newPlace);

                            var vertical45_2 = Global.UIDoc.Document.GetElement(elemIds.ToList()[0]) as MEPCurve;

                            Connector conHor1 = Common.GetConnectorValid(mepCurve);
                            if (conHor1 != null && !conHor1.IsConnected)
                            {
                                conHor1 = Common.GetConnectorNearest(conHor1.Origin, mepCurve, out Connector conHor2);

                                var level = Global.UIDoc.Document.GetElement(fastVerticalFrm.LevelId) as Level;
                                if (level == null)
                                    continue;

                                XYZ pointOffsetToLevel = Common.GetPointOffsetFromLevel(level, conHor1.Origin, fastVerticalFrm.OffSet);

                                if (pointOffsetToLevel.IsAlmostEqualTo(conHor1.Origin))
                                    continue;

                                Line line = Line.CreateBound(conHor1.Origin, pointOffsetToLevel);
                                (vertical45_2.Location as LocationCurve).Curve = line;

                                Common.ConnectPipeVerticalElbow45(Global.UIDoc.Document, mepCurve, vertical45_2, true);
                            }
                        }
                    }
                    else if (fastVerticalFrm.Siphon)
                    {
                        var end = new XYZ(p.X, p.Y, maxZ);

                        var elemIds = ElementTransformUtils.CopyElement(
                            Global.UIDoc.Document, mepCurve.Id, newPlace);

                        var vertical = Global.UIDoc.Document.GetElement(elemIds.ToList()[0]) as MEPCurve;

                        (vertical.Location as LocationCurve).Curve = Line.CreateBound(p, end);

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
                                var lv = Line.CreateBound(p, end);

                                ElementTransformUtils.RotateElement(Global.UIDoc.Document, vertical.Id, lv, a);
                            }
                        }

                        CreateSiphon(mepCurve, vertical, fastVerticalFrm.SiphonId);
                    }
                }
                t.Commit();
            }
            catch (Exception ex)
            {
                string mess = ex.Message;
            }

            if (t.HasStarted())
            {
                t.RollBack();
            }

            return Result.Succeeded;
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

        private static bool CreateElbowPipeIron(MEPCurve pipeHorizontal, Level level, double offsetFromLevel)
        {
            Pipe pipeVertical = null;
            double slope = Math.Round((double)Common.GetValueParameterByBuilt(pipeHorizontal, BuiltInParameter.RBS_PIPE_SLOPE), 5);

            if (Common.IsEqual(slope, 0))
            {
                Connector conHor1 = Common.GetConnectorValid(pipeHorizontal);
                if (conHor1 != null && !conHor1.IsConnected)
                {
                    conHor1 = Common.GetConnectorNearest(conHor1.Origin, pipeHorizontal, out Connector conHor2);

                    XYZ pointOffsetToLevel = Common.GetPointOffsetFromLevel(level, conHor1.Origin, offsetFromLevel);

                    if (pointOffsetToLevel.IsAlmostEqualTo(conHor1.Origin))
                        return false;

                    pipeVertical = Pipe.Create(Global.UIDoc.Document, pipeHorizontal.GetTypeId(), pipeHorizontal.ReferenceLevel.Id, conHor1, pointOffsetToLevel);
                    if (pipeVertical == null)
                        return false; ;
                }
            }
            else
            {
                XYZ pointSt = ((LocationCurve)pipeHorizontal.Location).Curve.GetEndPoint(0);

                Connector conSt = Common.GetConnectorNearest(pointSt, pipeHorizontal, out Connector conEnd);

                XYZ pointOffsetToLevelSt = Common.GetPointOffsetFromLevel(level, conSt.Origin, offsetFromLevel);

                XYZ pointOffsetToLevelEnd = Common.GetPointOffsetFromLevel(level, conEnd.Origin, offsetFromLevel);

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

                    pointOffsetToLevelSt = Common.GetPointOffsetFromLevel(level, retval.Origin, offsetFromLevel);
                    if (pointOffsetToLevelSt.IsAlmostEqualTo(retval.Origin))
                        return false;

                    pipeVertical = Pipe.Create(Global.UIDoc.Document, pipeHorizontal.GetTypeId(), pipeHorizontal.ReferenceLevel.Id, retval, pointOffsetToLevelSt);
                    if (pipeVertical == null)
                        return false;
                }
                else
                {
                    Connector conHor1 = Common.GetConnectorValid(pipeHorizontal);
                    if (conHor1 != null && !conHor1.IsConnected)
                    {
                        conHor1 = Common.GetConnectorNearest(conHor1.Origin, pipeHorizontal, out Connector conHor2);

                        XYZ pointOffsetToLevel = Common.GetPointOffsetFromLevel(level, conHor1.Origin, offsetFromLevel);

                        if (pointOffsetToLevel.IsAlmostEqualTo(conHor1.Origin))
                            return false;

                        pipeVertical = Pipe.Create(Global.UIDoc.Document, pipeHorizontal.GetTypeId(), pipeHorizontal.ReferenceLevel.Id, conHor1, pointOffsetToLevel);
                        if (pipeVertical == null)
                            return false;
                    }
                }
            }

            Common.GetConnectorClosedTo(pipeHorizontal.ConnectorManager, pipeVertical.ConnectorManager, out Connector con1, out Connector con2);
            if (con1 != null && con2 != null && con1.IsConnectedTo(con2))
                con1.DisconnectFrom(con2);
            Common.ConnectPipeVerticalElbow45(Global.UIDoc.Document, pipeHorizontal, pipeVertical, true);

            return true;
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