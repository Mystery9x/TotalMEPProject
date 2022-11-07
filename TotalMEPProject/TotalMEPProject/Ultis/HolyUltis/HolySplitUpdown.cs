using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;

//using MEPGenerator.Lisence;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;
using System.Windows.Forms;
using Autodesk.Revit.Creation;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI.Events;

namespace TotalMEPProject.Ultis.HolyUltis
{
    public class HolySplitUpdown
    {
        public UpDownMode m_UpDownMode = UpDownMode.UpAndDown;
        public MEPCurve _MEPCurve01 = null;
        public XYZ _Start01 = null;
        public XYZ _End01 = null;

        public MEPCurve _MEPCurve02 = null;
        public XYZ _Start02 = null;
        public XYZ _End02 = null;

        public List<XYZ> _Points = new List<XYZ>();

        public bool _IsChange = false;
        public bool IsRunning = false;

        public List<Element> _ElementAddeds = new List<Element>();

        private MEPCurve _MEPCurveChange01 = null;
        public XYZ _StartChange01 = null;
        public XYZ _EndChange01 = null;

        private MEPCurve _MEPCurveChange02 = null;
        public XYZ _StartChange02 = null;
        public XYZ _EndChange02 = null;

        private MEPCurve _MEPCurveSplit01 = null;
        private MEPCurve _MEPCurveSplit02 = null;

        private MEPCurve _MEPCurve_Vertical01 = null;
        private MEPCurve _MEPCurve_Vertical02 = null;

        private Element _Elbow1 = null;
        private Element _Elbow2 = null;

        private Element _Elbow3 = null;
        private Element _Elbow4 = null;

        public Result OK()
        {
            Global.RVTApp.DocumentChanged -= new EventHandler<DocumentChangedEventArgs>(OnDocumentChanged);

            return Result.Succeeded;
        }

        public Result Run()
        {
            if (_MEPCurve01 == null)
                return Result.Cancelled;

            if (_Points.Count == 0)
                return Result.Cancelled;

            if (m_UpDownMode == UpDownMode.UpOrDown)
            {
                if (_IsChange == false)
                {
                    return Run_Up_Down();
                }
                else
                {
                    return Change_Up_Down();
                }
            }
            else
            {
                if (_MEPCurve01 != null && _MEPCurve02 == null)
                    return Run_UpAndDown01();
                else
                    return Run_UpAndDown02();
            }
        }

        private ElementId Split(MEPCurve mep, XYZ p)
        {
            ElementId splitMEPId = ElementId.InvalidElementId;

            try
            {
                if (mep as Pipe != null)
                {
                    splitMEPId = PlumbingUtils.BreakCurve(Global.UIDoc.Document, mep.Id, p);
                }
                else if (mep as Duct != null)
                {
                    splitMEPId = MechanicalUtils.BreakCurve(Global.UIDoc.Document, mep.Id, p);
                }
                else if (mep as CableTray != null)
                {
                }
                else if (mep as Conduit != null)
                {
                }

                return splitMEPId;
            }
            catch (System.Exception ex)
            {
                return splitMEPId;
            }
        }

        public Result Change_Up_Down()
        {
            Transaction tran = new Transaction(Global.UIDoc.Document, "Holy");
            tran.Start();

            try
            {
                double offset = Math.Abs(App.m_HolyUpDownForm.Distance) * Common.mmToFT;
                if (App.m_HolyUpDownForm.Distance < 0)
                {
                    offset *= -1;
                }

                var pSplit = _Points[0];

                (_MEPCurveChange01.Location as LocationCurve).Curve = Line.CreateBound(_StartChange01, _EndChange01);

                List<ElementId> deletes = new List<ElementId>();

                deletes.Add(_MEPCurve_Vertical01.Id);
                deletes.Add(_Elbow1.Id);
                deletes.Add(_Elbow2.Id);

                //Thay doi cao do cua mep
                var lineChange = (_MEPCurveChange01.Location as LocationCurve).Curve.Clone() as Line;

                var result = Process_Up_Down(lineChange, offset, pSplit);

                if (result == Result.Succeeded)
                {
                    foreach (ElementId id in deletes)
                    {
                        Global.UIDoc.Document.Delete(id);
                    }

                    tran.Commit();
                    return Result.Succeeded;
                }
                else
                {
                    tran.RollBack();
                    return Result.Cancelled;
                }
            }
            catch (System.Exception ex)
            {
                tran.RollBack();
                return Result.Cancelled;
            }
        }

        public Result Run_UpAndDown02()
        {
            Transaction tran = new Transaction(Global.UIDoc.Document, "Holy");
            tran.Start();

            IsRunning = true;

            var offset = Math.Abs(App.m_HolyUpDownForm.Distance) * Common.mmToFT;

            try
            {
                List<ElementId> deletes = new List<ElementId>();
                if (_IsChange)
                {
                    deletes.Add(_MEPCurve_Vertical01.Id);
                    deletes.Add(_MEPCurve_Vertical02.Id);
                    deletes.Add(_Elbow1.Id);
                    deletes.Add(_Elbow2.Id);
                    deletes.Add(_Elbow3.Id);
                    deletes.Add(_Elbow4.Id);
                }

                var splitPoint1 = _Points[0];
                var splitPoint2 = _Points[1];

                if (_MEPCurve01 != null && _MEPCurve02 != null && _MEPCurveSplit01 == null && _MEPCurveSplit02 == null)
                {
                    var splitMEPId1 = Split(_MEPCurve01, splitPoint1);
                    if (splitMEPId1 == ElementId.InvalidElementId)
                    {
                        tran.RollBack();
                        return Result.Cancelled;
                    }

                    _MEPCurveSplit01 = Global.UIDoc.Document.GetElement(splitMEPId1) as MEPCurve;

                    var splitMEPId2 = Split(_MEPCurve02, splitPoint2);
                    if (splitMEPId2 == ElementId.InvalidElementId)
                    {
                        tran.RollBack();
                        return Result.Cancelled;
                    }

                    _MEPCurveSplit02 = Global.UIDoc.Document.GetElement(splitMEPId2) as MEPCurve;
                }

                if (_MEPCurveSplit01 == null || _MEPCurveSplit02 == null)
                {
                    tran.RollBack();
                    return Result.Cancelled;
                }

                if (App.m_HolyUpDownForm.Distance < 0)
                {
                    offset *= -1;
                }

                var result = Process_UpAndDown02(offset);

                if (result == Result.Succeeded)
                {
                    foreach (ElementId id in deletes)
                    {
                        Global.UIDoc.Document.Delete(id);
                    }
                    tran.Commit();

                    App._HolyUpdown._IsChange = true;

                    return Result.Succeeded;
                }
                else
                {
                    tran.RollBack();

                    if (_IsChange == false)
                    {
                        _MEPCurveSplit01 = null;
                        _MEPCurveSplit02 = null;
                    }

                    return Result.Cancelled;
                }
            }
            catch (System.Exception ex)
            {
                string mess = ex.Message;

                if (tran.HasStarted())
                {
                    tran.RollBack();
                }

                return Result.Cancelled;
            }
        }

        private void Detect()
        {
            var curve1 = (_MEPCurve01.Location as LocationCurve).Curve;
            var center1 = curve1.Evaluate((curve1.GetEndParameter(0) + curve1.GetEndParameter(1) / 2), false);

            var curve2 = (_MEPCurve02.Location as LocationCurve).Curve;
            var center2 = curve2.Evaluate((curve2.GetEndParameter(0) + curve2.GetEndParameter(1) / 2), false);

            var curve3 = (_MEPCurveSplit01.Location as LocationCurve).Curve;
            var center3 = curve3.Evaluate((curve3.GetEndParameter(0) + curve3.GetEndParameter(1) / 2), false);

            var curve4 = (_MEPCurveSplit02.Location as LocationCurve).Curve;
            var center4 = curve4.Evaluate((curve4.GetEndParameter(0) + curve4.GetEndParameter(1) / 2), false);

            if (center1.DistanceTo(_Points[1]) < center3.DistanceTo(_Points[1]))
            {
                _MEPCurveChange01 = _MEPCurve01;
                _MEPCurve01 = _MEPCurveSplit01;
            }
            else
            {
                _MEPCurveChange01 = _MEPCurveSplit01;
            }

            if (center2.DistanceTo(_Points[0]) < center4.DistanceTo(_Points[0]))
            {
                _MEPCurveChange02 = _MEPCurve02;

                _MEPCurve02 = _MEPCurveSplit02;
            }
            else
            {
                _MEPCurveChange02 = _MEPCurveSplit02;
            }
        }

        private Result Process_UpAndDown02(double offset)
        {
            if (_IsChange == false)
                Detect();

            if (_MEPCurveChange01 == null || _MEPCurveChange02 == null)
                return Result.Cancelled;

            var splitPoint1 = _Points[0];
            var splitPoint2 = _Points[1];

            XYZ newPoint1 = null;
            XYZ newPoint2 = null;

            var curve1 = (_MEPCurveChange01.Location as LocationCurve).Curve;
            var curve2 = (_MEPCurveChange02.Location as LocationCurve).Curve;

            if (_IsChange == false)
            {
                _StartChange01 = curve1.GetEndPoint(0);
                _EndChange01 = curve1.GetEndPoint(1);

                _StartChange02 = curve2.GetEndPoint(0);
                _EndChange02 = curve2.GetEndPoint(1);
            }
            if (App.m_HolyUpDownForm.Elbow90)
            {
                var pStart1 = new XYZ(_StartChange01.X, _StartChange01.Y, _StartChange01.Z + offset);
                var pEnd1 = new XYZ(_EndChange01.X, _EndChange01.Y, _EndChange01.Z + offset);

                (_MEPCurveChange01.Location as LocationCurve).Curve = Line.CreateBound(pStart1, pEnd1);

                var pStart2 = new XYZ(_StartChange02.X, _StartChange02.Y, _StartChange02.Z + offset);
                var pEnd2 = new XYZ(_EndChange02.X, _EndChange02.Y, _EndChange02.Z + offset);

                (_MEPCurveChange02.Location as LocationCurve).Curve = Line.CreateBound(pStart2, pEnd2);

                newPoint1 = new XYZ(splitPoint1.X, splitPoint1.Y, splitPoint1.Z + offset);
                newPoint2 = new XYZ(splitPoint2.X, splitPoint2.Y, splitPoint2.Z + offset);
            }
            else
            {
                var lineChange01 = Line.CreateBound(_StartChange01, _EndChange01);
                if (lineChange01.Length <= Math.Abs(offset))
                {
                    MessageBox.Show("Offset is invalid.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return Result.Cancelled;
                }

                if (splitPoint1.DistanceTo(_StartChange01) < splitPoint1.DistanceTo(_EndChange01))
                {
                    var newPoint = lineChange01.Evaluate(Math.Abs(offset), false);
                    XYZ pOther = lineChange01.GetEndPoint(1);
                    lineChange01 = Line.CreateBound(newPoint, pOther);
                    newPoint1 = new XYZ(newPoint.X, newPoint.Y, newPoint.Z + offset);
                }
                else
                {
                    var newPoint = lineChange01.Evaluate(lineChange01.Length - Math.Abs(offset), false);
                    XYZ pOther = lineChange01.GetEndPoint(0);
                    lineChange01 = Line.CreateBound(pOther, newPoint);
                    newPoint1 = new XYZ(newPoint.X, newPoint.Y, newPoint.Z + offset);
                }

                var pStart1 = new XYZ(lineChange01.GetEndPoint(0).X, lineChange01.GetEndPoint(0).Y, lineChange01.GetEndPoint(0).Z + offset);
                var pEnd1 = new XYZ(lineChange01.GetEndPoint(1).X, lineChange01.GetEndPoint(1).Y, lineChange01.GetEndPoint(1).Z + offset);

                (_MEPCurveChange01.Location as LocationCurve).Curve = Line.CreateBound(pStart1, pEnd1);

                var lineChange02 = Line.CreateBound(_StartChange02, _EndChange02);
                if (lineChange02.Length <= Math.Abs(offset))
                {
                    MessageBox.Show("Offset is invalid.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return Result.Cancelled;
                }

                if (splitPoint2.DistanceTo(_StartChange02) < splitPoint2.DistanceTo(_EndChange02))
                {
                    var newPoint = lineChange02.Evaluate(Math.Abs(offset), false);
                    XYZ pOther = lineChange02.GetEndPoint(1);
                    lineChange02 = Line.CreateBound(newPoint, pOther);
                    newPoint2 = new XYZ(newPoint.X, newPoint.Y, newPoint.Z + offset);
                }
                else
                {
                    var newPoint = lineChange02.Evaluate(lineChange02.Length - Math.Abs(offset), false);
                    XYZ pOther = lineChange02.GetEndPoint(0);
                    lineChange02 = Line.CreateBound(pOther, newPoint);
                    newPoint2 = new XYZ(newPoint.X, newPoint.Y, newPoint.Z + offset);
                }
                var pStart2 = new XYZ(lineChange02.GetEndPoint(0).X, lineChange02.GetEndPoint(0).Y, lineChange02.GetEndPoint(0).Z + offset);
                var pEnd2 = new XYZ(lineChange02.GetEndPoint(1).X, lineChange02.GetEndPoint(1).Y, lineChange02.GetEndPoint(1).Z + offset);

                (_MEPCurveChange02.Location as LocationCurve).Curve = Line.CreateBound(pStart2, pEnd2);
            }

            _MEPCurve_Vertical01 = MEPUtilscs.CC(_MEPCurve01, Line.CreateBound(splitPoint1, newPoint1));
            _MEPCurve_Vertical02 = MEPUtilscs.CC(_MEPCurve01, Line.CreateBound(splitPoint2, newPoint2));

            _Elbow1 = MEPUtilscs.CE(_MEPCurveChange01, _MEPCurve_Vertical01, newPoint1);
            _Elbow2 = MEPUtilscs.CE(_MEPCurveChange02, _MEPCurve_Vertical02, newPoint2);

            _Elbow3 = MEPUtilscs.CE(_MEPCurve01, _MEPCurve_Vertical01, splitPoint1);
            _Elbow4 = MEPUtilscs.CE(_MEPCurve02, _MEPCurve_Vertical02, splitPoint2);

            return Result.Succeeded;
        }

        public Result Run_UpAndDown01()
        {
            Transaction tran = new Transaction(Global.UIDoc.Document, "Holy");
            tran.Start();

            IsRunning = true;

            var offset = Math.Abs(App.m_HolyUpDownForm.Distance) * Common.mmToFT;

            try
            {
                var splitPoint1 = _Points[0];
                var splitPoint2 = _Points[1];
                if (_Start01.DistanceTo(splitPoint1) > _Start01.DistanceTo(splitPoint2))
                {
                    splitPoint1 = _Points[1];
                    splitPoint2 = _Points[0];
                }

                if (_MEPCurveSplit01 == null)
                {
                    var splitMEPId1 = Split(_MEPCurve01, splitPoint1);
                    if (splitMEPId1 == ElementId.InvalidElementId)
                    {
                        tran.RollBack();
                        return Result.Cancelled;
                    }

                    _MEPCurveSplit01 = Global.UIDoc.Document.GetElement(splitMEPId1) as MEPCurve;

                    var splitMEPId2 = Split(_MEPCurve01, splitPoint2);
                    if (splitMEPId2 != ElementId.InvalidElementId)
                    {
                        _MEPCurveChange01 = Global.UIDoc.Document.GetElement(splitMEPId2) as MEPCurve;
                    }
                    else
                    {
                        splitMEPId2 = Split(_MEPCurveSplit01, splitPoint2);
                        if (splitMEPId2 != ElementId.InvalidElementId)
                        {
                            _MEPCurveChange01 = Global.UIDoc.Document.GetElement(splitMEPId2) as MEPCurve;
                        }
                    }
                }

                if (_MEPCurveSplit01 == null || _MEPCurveChange01 == null)
                {
                    tran.RollBack();
                    return Result.Cancelled;
                }

                if (App.m_HolyUpDownForm.Distance < 0)
                {
                    offset *= -1;
                }

                Line lineChange = Line.CreateBound(splitPoint1, splitPoint2);

                List<ElementId> deletes = new List<ElementId>();
                if (_IsChange == false)
                {
                    _StartChange01 = lineChange.GetEndPoint(0);
                    _EndChange01 = lineChange.GetEndPoint(1);
                }
                else
                {
                    (_MEPCurveChange01.Location as LocationCurve).Curve = Line.CreateBound(_StartChange01, _EndChange01);
                    deletes.Add(_MEPCurve_Vertical01.Id);
                    deletes.Add(_MEPCurve_Vertical02.Id);
                    deletes.Add(_Elbow1.Id);
                    deletes.Add(_Elbow2.Id);
                    deletes.Add(_Elbow3.Id);
                    deletes.Add(_Elbow4.Id);
                }

                var result = Process_UpAndDown01(lineChange, splitPoint1, splitPoint2, offset);

                if (result == Result.Succeeded)
                {
                    foreach (ElementId id in deletes)
                    {
                        Global.UIDoc.Document.Delete(id);
                    }

                    tran.Commit();

                    App._HolyUpdown._IsChange = true;

                    return Result.Succeeded;
                }
                else
                {
                    tran.RollBack();

                    if (_IsChange == false)
                    {
                        _MEPCurveSplit01 = null;
                        _MEPCurveChange01 = null;
                    }

                    return Result.Cancelled;
                }
            }
            catch (System.Exception ex)
            {
                string mess = ex.Message;

                if (tran.HasStarted())
                {
                    tran.RollBack();
                }

                return Result.Cancelled;
            }
        }

        public Result Run_Up_Down()
        {
            Transaction tran = new Transaction(Global.UIDoc.Document, "Holy");
            tran.Start();

            IsRunning = true;
            var offset = Math.Abs(App.m_HolyUpDownForm.Distance) * Common.mmToFT;

            try
            {
                bool isSlope = false;
                var para = _MEPCurve01.LookupParameter("Slope");
                if (para != null)
                {
                    var slope = para.AsDouble();
                    if (slope != 0)
                    {
                        isSlope = true;
                    }
                }

                if (_MEPCurveSplit01 == null)
                {
                    var splitMEPId = Split(_MEPCurve01, _Points[0]);
                    if (splitMEPId == ElementId.InvalidElementId)
                    {
                        tran.RollBack();
                        return Result.Cancelled;
                    }

                    _MEPCurveSplit01 = Global.UIDoc.Document.GetElement(splitMEPId) as MEPCurve;
                }

                var curve1 = _MEPCurve01.GetCurve();
                var curve2 = _MEPCurveSplit01.GetCurve();

                var p0 = _Start01;
                var p1 = _End01;

                if (App.m_HolyUpDownForm.Distance < 0)
                {
                    offset *= -1;
                }

                var pSplit = _Points[0];

                if (isSlope)
                {
                    if (p0.Z > p1.Z)
                    {
                        if (App.m_HolyUpDownForm.Distance < 0)
                        {
                            _MEPCurveChange01 = _MEPCurve01;
                        }
                        else
                        {
                            _MEPCurveChange01 = _MEPCurveSplit01;
                        }
                    }
                    else
                    {
                        if (App.m_HolyUpDownForm.Distance < 0)
                        {
                            _MEPCurveChange01 = _MEPCurveSplit01;
                        }
                        else
                        {
                            _MEPCurveChange01 = _MEPCurve01;
                        }
                    }
                }
                else if (_Points.Count == 2)
                {
                    var pSide = _Points[1];

                    var line = Line.CreateBound(p0, p1);
                    var proSide = line.Project(pSide);
                    var proSplit = line.Project(_Points[0]);

                    if (proSide.Parameter < proSplit.Parameter)
                    {
                        _MEPCurveChange01 = _MEPCurveSplit01;
                    }
                    else
                    {
                        _MEPCurveChange01 = _MEPCurve01;
                    }
                }
                else
                    return Result.Cancelled;

                if (_MEPCurveChange01 == null)
                {
                    return Result.Cancelled;
                }

                var lineChange = (_MEPCurveChange01.Location as LocationCurve).Curve.Clone() as Line;

                lineChange = Line.CreateBound(lineChange.GetEndPoint(0), lineChange.GetEndPoint(1));

                _StartChange01 = lineChange.GetEndPoint(0);
                _EndChange01 = lineChange.GetEndPoint(1);

                var result = Process_Up_Down(lineChange, offset, pSplit);

                if (result == Result.Succeeded)
                {
                    tran.Commit();

                    App._HolyUpdown._IsChange = true;

                    return Result.Succeeded;
                }
                else
                {
                    tran.RollBack();

                    return Result.Cancelled;
                }
            }
            catch (System.Exception ex)
            {
                string mess = ex.Message;

                if (tran.HasStarted())
                {
                    tran.RollBack();
                }

                return Result.Cancelled;
            }
        }

        private Result Process_UpAndDown01(Line lineChange, XYZ splitPoint1, XYZ splitPoint2, double offset)
        {
            XYZ newPoint1 = null;
            XYZ newPoint2 = null;

            if (App.m_HolyUpDownForm.Elbow90)
            {
                newPoint1 = new XYZ(splitPoint1.X, splitPoint1.Y, splitPoint1.Z + offset);
                newPoint2 = new XYZ(splitPoint2.X, splitPoint2.Y, splitPoint2.Z + offset);
            }
            else
            {
                if (lineChange.Length <= Math.Abs(offset))
                {
                    MessageBox.Show("Offset is invalid.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return Result.Cancelled;
                }

                XYZ newPoint = null;
                newPoint = lineChange.Evaluate(Math.Abs(offset), false);
                XYZ pOther = lineChange.GetEndPoint(1);
                lineChange = Line.CreateBound(newPoint, pOther);
                newPoint1 = new XYZ(newPoint.X, newPoint.Y, newPoint.Z + offset);

                if (lineChange.Length <= Math.Abs(offset))
                {
                    MessageBox.Show("Offset is invalid.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return Result.Cancelled;
                }

                newPoint = lineChange.Evaluate(lineChange.Length - Math.Abs(offset), false);
                pOther = lineChange.GetEndPoint(0);
                lineChange = Line.CreateBound(pOther, newPoint);
                newPoint2 = new XYZ(newPoint.X, newPoint.Y, newPoint.Z + offset);
            }

            var pStart = new XYZ(lineChange.GetEndPoint(0).X, lineChange.GetEndPoint(0).Y, lineChange.GetEndPoint(0).Z + offset);
            var pEnd = new XYZ(lineChange.GetEndPoint(1).X, lineChange.GetEndPoint(1).Y, lineChange.GetEndPoint(1).Z + offset);

            (_MEPCurveChange01.Location as LocationCurve).Curve = Line.CreateBound(pStart, pEnd);

            _MEPCurve_Vertical01 = MEPUtilscs.CC(_MEPCurve01, Line.CreateBound(splitPoint1, newPoint1));
            _MEPCurve_Vertical02 = MEPUtilscs.CC(_MEPCurve01, Line.CreateBound(splitPoint2, newPoint2));

            _Elbow1 = MEPUtilscs.CE(_MEPCurveChange01, _MEPCurve_Vertical01, newPoint1);
            _Elbow2 = MEPUtilscs.CE(_MEPCurveChange01, _MEPCurve_Vertical02, newPoint2);

            _Elbow3 = MEPUtilscs.CE(_MEPCurveSplit01, _MEPCurve_Vertical01, splitPoint1);
            _Elbow4 = MEPUtilscs.CE(_MEPCurve01, _MEPCurve_Vertical02, splitPoint2);

            return Result.Succeeded;
        }

        private Result Process_Up_Down(Line lineChange, double offset, XYZ pSplit)
        {
            XYZ newPoint = null;
            if (App.m_HolyUpDownForm.Elbow90)
            {
                newPoint = new XYZ(pSplit.X, pSplit.Y, pSplit.Z + offset);
            }
            else
            {
                if (_MEPCurveChange01.Id == _MEPCurveSplit01.Id)
                {
                    newPoint = lineChange.Evaluate(lineChange.Length - Math.Abs(offset), false);
                    XYZ pOther = lineChange.GetEndPoint(0);
                    lineChange = Line.CreateBound(pOther, newPoint);
                }
                else
                {
                    newPoint = lineChange.Evaluate(Math.Abs(offset), false);
                    XYZ pOther = lineChange.GetEndPoint(1);
                    lineChange = Line.CreateBound(newPoint, pOther);
                }

                newPoint = new XYZ(newPoint.X, newPoint.Y, newPoint.Z + offset);
            }

            var pStart = new XYZ(lineChange.GetEndPoint(0).X, lineChange.GetEndPoint(0).Y, lineChange.GetEndPoint(0).Z + offset);
            var pEnd = new XYZ(lineChange.GetEndPoint(1).X, lineChange.GetEndPoint(1).Y, lineChange.GetEndPoint(1).Z + offset);

            //Change old mep
            (_MEPCurveChange01.Location as LocationCurve).Curve = Line.CreateBound(pStart, pEnd);

            _MEPCurve_Vertical01 = MEPUtilscs.CC(_MEPCurve01, Line.CreateBound(pSplit, newPoint));

            if (_MEPCurve_Vertical01 != null)
            {
                if (_MEPCurveChange01.Id != _MEPCurve01.Id)
                {
                    _Elbow1 = MEPUtilscs.CE(_MEPCurve01, _MEPCurve_Vertical01, pSplit);
                    _Elbow2 = MEPUtilscs.CE(_MEPCurveChange01, _MEPCurve_Vertical01, newPoint);
                }
                else
                {
                    _Elbow1 = MEPUtilscs.CE(_MEPCurveChange01, _MEPCurve_Vertical01, newPoint);
                    _Elbow2 = MEPUtilscs.CE(_MEPCurveSplit01, _MEPCurve_Vertical01, pSplit);
                }
            }

            return Result.Succeeded;
        }

        public void OnDocumentChanged(object sender, DocumentChangedEventArgs args)
        {
            if (IsRunning)
            {
                var doc = args.GetDocument();
                UndoOperation ope = args.Operation;
                string tranName = "";
                foreach (string name in args.GetTransactionNames())
                {
                    tranName += name;
                }

                if (ope == UndoOperation.TransactionUndone)
                {
                }

                foreach (ElementId newId in args.GetAddedElementIds())
                {
                    var element = Global.UIDoc.Document.GetElement(newId) as Element;
                    if (element as MEPCurve != null || element as FamilyInstance != null)
                    {
                        if (_ElementAddeds.Contains(element) == false)
                            _ElementAddeds.Add(element);
                    }
                }
            }
        }
    }
}