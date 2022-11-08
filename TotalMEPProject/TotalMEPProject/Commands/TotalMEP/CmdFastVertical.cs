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
            FastVerticalForm form = new FastVerticalForm(_levels);
            if (form.ShowDialog() != DialogResult.OK)
                return Result.Cancelled;

            //Select pipes
            List<MEPCurve> mepCurves = SelectMEPCurves(form.GetMEPType);

            if (mepCurves == null || mepCurves.Count == 0)
                return Result.Cancelled;

            Transaction t = new Transaction(Global.UIDoc.Document, "a");
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

                var offset = form.OffSet * Common.mmToFT;

                //////////////////////////////////////////////////////////////////////////
                bool isUp = true;

                //////////////////////////////////////////////////////////////////////////

                double maxZ = 0;
                if (form.ByLevel)
                {
                    var level = Global.UIDoc.Document.GetElement(form.LevelId) as Level;
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

                if (form.Elbow90)
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
                else if (form.Elbow45)
                {
                    //Tao mot pipe 45
                    var elemIds = ElementTransformUtils.CopyElement(
                        Global.UIDoc.Document, mepCurve.Id, newPlace);

                    var vertical45_2 = Global.UIDoc.Document.GetElement(elemIds.ToList()[0]) as MEPCurve;

                    Connector conHor1 = Common.GetConnectorValid(mepCurve);
                    if (conHor1 != null && !conHor1.IsConnected)
                    {
                        conHor1 = Common.GetConnectorNearest(conHor1.Origin, mepCurve, out Connector conHor2);

                        var level = Global.UIDoc.Document.GetElement(form.LevelId) as Level;
                        if (level == null)
                            continue;

                        XYZ pointOffsetToLevel = Common.GetPointOffsetFromLevel(level, conHor1.Origin, form.OffSet);

                        if (pointOffsetToLevel.IsAlmostEqualTo(conHor1.Origin))
                            continue;

                        Line line = Line.CreateBound(conHor1.Origin, pointOffsetToLevel);
                        (vertical45_2.Location as LocationCurve).Curve = line;

                        Common.ConnectPipeVerticalElbow45(Global.UIDoc.Document, mepCurve, vertical45_2, true);
                    }
                }
                else if (form.Siphon)
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

                    CreateSiphon(mepCurve, vertical, form.SiphonId);
                }
            }
            t.Commit();

            return Result.Succeeded;
        }

        private FamilyInstance CreateSiphon(MEPCurve mepCurve, MEPCurve vertical_pipe, ElementId symbolId)
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

        private void SetCurveNormal(CableTray vertical, CableTray current)
        {
            var curve = current.GetCurve() as Line;
            if (curve == null)
                return;

            var v = curve.Direction;

            (vertical as CableTray).CurveNormal = v;
        }

        public List<MEPCurve> SelectMEPCurves(Type type)
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

        private FamilyInstance CreateElbow(MEPCurve pipe1, MEPCurve pipe2)
        {
            try
            {
                List<Connector> connectors = Common.GetConnectionNearest(pipe1, pipe2);
                if (connectors != null && connectors.Count == 2)
                {
                    var elbow = Global.UIDoc.Document.Create.NewElbowFitting(connectors[0], connectors[1]);

                    return elbow;

                    //                             //set
                    //                             var con = Utils.GetConnectorClosestTo(elbow, end);
                    //
                    //                             var d = height - con.Origin.Z;
                    //                             end = new XYZ(con.Origin.X, con.Origin.Y, con.Origin.Z + d);
                    //                             (vertical.Location as LocationCurve).Curve = Line.CreateBound(con.Origin, end);
                }
            }
            catch (System.Exception ex)
            {
            }
            return null;
        }

        private int GetAt(MEPCurve pipe)
        {
            var curve = (pipe.Location as LocationCurve).Curve;

            var p0 = curve.GetEndPoint(0);
            var p1 = curve.GetEndPoint(1);

            bool atStart = false;
            //             var con = Utils.GetConnectorClosestTo(pipe, p0);
            //             if (con.AllRefs.Size != 0)
            //                 atStart = true;
            //             else
            atStart = CheckIntersection(p0, pipe);

            //             con = Utils.GetConnectorClosestTo(pipe, p1);
            //
            bool atEnd = false;
            //             if (con.AllRefs.Size != 0)
            //                 atEnd = true;
            //             else
            atEnd = CheckIntersection(p1, pipe);

            if (atStart == true && atEnd == false)
                return 1;
            else if (atStart == false && atEnd == true)
                return 0;
            else if (atStart == false && atEnd == false) // Ko intersection thi uu tien tai end
                return 1;

            return -1;
        }

        private bool CheckIntersection(XYZ p, MEPCurve pipe)
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