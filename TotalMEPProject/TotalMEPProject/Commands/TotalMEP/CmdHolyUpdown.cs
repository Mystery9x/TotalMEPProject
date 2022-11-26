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
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using TotalMEPProject.UI;
using TotalMEPProject.Ultis;
using TotalMEPProject.Ultis.HolyUltis;
using TotalMEPProject.Ultis.StorageUtility;

namespace TotalMEPProject.Commands.TotalMEP
{
    [Transaction(TransactionMode.Manual)]
    public class CmdHolyUpdown : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Global.UIApp = commandData.Application;
            Global.RVTApp = commandData.Application.Application;
            Global.UIDoc = commandData.Application.ActiveUIDocument;
            Global.AppCreation = commandData.Application.Application.Create;

            if (App.ShowHolyUpDownForm() == false)
            {
                return Result.Cancelled;
            };
            return Result.Succeeded;
        }

        public static List<Data> _Datas = new List<Data>();

        public static Result Run_PickObjects()
        {
            try
            {
                var meps = GetSelectedMEP();
                if (meps == null || meps.Count == 0)
                    return Result.Cancelled;

                _Datas.Clear();

                var schema = Autodesk.Revit.DB.ExtensibleStorage.Schema.Lookup(StorageUtility.m_MEP_HoLyUpDown_Guild);
                if (schema != null)
                {
                    object valueEntity = StorageUtility.GetValue(Global.UIDoc.Document.ProjectInformation, schema, StorageUtility.m_MEP_HoLyUpDown, typeof(string));

                    if (valueEntity != null)
                    {
                        var str = valueEntity as string;

                        ReadData(str);
                    }
                }

                if (_Datas.Count != 0)
                {
                    foreach (MEPCurve mepCurve in meps)
                    {
                        var find = _Datas.Find(item => item._MEPCurve.Id == mepCurve.Id);
                        if (find == null)
                        {
                            _Datas.Clear();
                            break;
                        }
                    }
                }

                double offsetMain = UnitUtils.ConvertToInternalUnits(Math.Abs(App.m_HolyUpDownForm.Distance), DisplayUnitType.DUT_MILLIMETERS);// Math.Abs(App.m_HolyUpDownForm.Distance) * Common.mmToFT;
                if (App.m_HolyUpDownForm.Distance < 0)
                {
                    offsetMain *= -1;
                }

                Transaction tran = new Transaction(Global.UIDoc.Document, "Holy");
                tran.Start();

                Dictionary<MEPCurve, List<FamilyInstance>> data = new Dictionary<MEPCurve, List<FamilyInstance>>();
                Dictionary<MEPCurve, List<FamilyInstance>> data_tap = new Dictionary<MEPCurve, List<FamilyInstance>>();

                foreach (MEPCurve mepCurve in meps)
                {
                    data.Add(mepCurve, new List<FamilyInstance>());
                    data_tap.Add(mepCurve, new List<FamilyInstance>());

                    //Get union
                    foreach (Connector c in mepCurve.ConnectorManager.Connectors)
                    {
                        foreach (Connector cfitting in c.AllRefs)
                        {
                            var fitting = cfitting.Owner as FamilyInstance;
                            if (fitting == null)
                                continue;

                            var location = (fitting.Location as LocationPoint).Point;

                            var mechanicalFitting = fitting.MEPModel as MechanicalFitting;

                            bool flag = false;
                            if (mechanicalFitting != null)
                            {
                                if (mechanicalFitting.PartType == PartType.Union)
                                {
                                    cfitting.DisconnectFrom(c);

                                    flag = true;

                                    data[mepCurve].Add(fitting);
                                }
                                else
                                {
                                    if (mechanicalFitting.PartType == PartType.TapPerpendicular || mechanicalFitting.PartType == PartType.TapAdjustable)
                                    {
                                        cfitting.DisconnectFrom(c);
                                        data_tap[mepCurve].Add(fitting);
                                    }
                                }
                            }
                            else
                            {
                                var family = fitting.Symbol.Family;
                                var partypePara = family.get_Parameter(BuiltInParameter.FAMILY_CONTENT_PART_TYPE);
                                if (partypePara != null)
                                {
                                    var type = partypePara.AsValueString();

                                    if (type.Contains("Union"))
                                    {
                                        cfitting.DisconnectFrom(c);

                                        flag = true;

                                        data[mepCurve].Add(fitting);
                                    }
                                }
                            }

                            if (flag)
                            {
                                var dataMep = _Datas.Find(item => item._MEPCurve.Id == mepCurve.Id);
                                if (dataMep == null)
                                {
                                    //Reset by location
                                    var curve = (mepCurve.Location as LocationCurve).Curve;
                                    var p0 = curve.GetEndPoint(0);
                                    var p1 = curve.GetEndPoint(1);

                                    if (p0.DistanceTo(location) < p1.DistanceTo(location))
                                    {
                                        var p0New = new XYZ();
                                        (mepCurve.Location as LocationCurve).Curve = Line.CreateBound(location, p1);
                                    }
                                    else
                                    {
                                        (mepCurve.Location as LocationCurve).Curve = Line.CreateBound(p0, location);
                                    }
                                }
                            }
                        }
                    }
                }
                if (data.Count != 0)
                {
                    //Add to list

                    Dictionary<MEPCurve, bool> data_exist = new Dictionary<MEPCurve, bool>();
                    foreach (KeyValuePair<MEPCurve, List<FamilyInstance>> keyPair in data)
                    {
                        var mepCurve = keyPair.Key;
                        var unions = keyPair.Value;

                        var dataMep = _Datas.Find(item => item._MEPCurve.Id == mepCurve.Id);

                        double offset = offsetMain;
                        if (dataMep == null)
                        {
                            dataMep = new Data(mepCurve);
                            var curve = (mepCurve.Location as LocationCurve).Curve;
                            dataMep._Start = curve.GetEndPoint(0);
                            dataMep._End = curve.GetEndPoint(1);
                            dataMep._OldOffset = offset;
                            _Datas.Add(dataMep);

                            data_exist.Add(mepCurve, false);
                        }
                        else
                        {
                            data_exist.Add(mepCurve, true);
                        }
                        if (dataMep._Start != null && dataMep._End != null)
                        {
                            if (App.m_HolyUpDownForm.OriginalElevation == false)
                            {
                                offset += dataMep._OldOffset;
                                dataMep._OldOffset = offset;
                            }
                            else
                                dataMep._OldOffset = offset;

                            var p0New = OffsetZ(dataMep._Start, offset);
                            var p1New = OffsetZ(dataMep._End, offset);

                            (mepCurve.Location as LocationCurve).Curve = Line.CreateBound(p0New, p1New);

                            //Move other mep connect to tap
                            if (data_tap.ContainsKey(mepCurve) == true)
                            {
                                foreach (FamilyInstance tap in data_tap[mepCurve])
                                {
                                    var p = (tap.Location as LocationPoint).Point;

                                    var pZ = OffsetZ(p, offset);

                                    MEPCurve other = null;
                                    foreach (Connector c in tap.MEPModel.ConnectorManager.Connectors)
                                    {
                                        foreach (Connector cmep in c.AllRefs)
                                        {
                                            var mep = cmep.Owner as MEPCurve;
                                            if (mep != null)
                                            {
                                                other = mep;
                                                break;
                                            }
                                        }

                                        if (other != null)
                                            break;
                                    }

                                    if (other != null)
                                    {
                                        Global.UIDoc.Document.Delete(tap.Id);

                                        try
                                        {
                                            var curve = (other.Location as LocationCurve).Curve;

                                            var p0 = curve.GetEndPoint(0);
                                            var p1 = curve.GetEndPoint(1);

                                            p0 = new XYZ(p0.X, p0.Y, p0New.Z);
                                            p1 = new XYZ(p1.X, p1.Y, p0New.Z);

                                            (other.Location as LocationCurve).Curve = Line.CreateBound(p0, p1);

                                            var con = GetConnectorClosestTo(other, pZ);
                                            var elbow = Global.UIDoc.Document.Create.NewTakeoffFitting(con, mepCurve);
                                        }
                                        catch (System.Exception ex)
                                        {
                                        }
                                    }
                                }
                            }
                        }
                    }

                    //Delete old element
                    foreach (KeyValuePair<MEPCurve, List<FamilyInstance>> keyPair in data)
                    {
                        var mepCurve = keyPair.Key;
                        var unions = keyPair.Value;

                        var dataMep = _Datas.Find(item => item._MEPCurve.Id == mepCurve.Id);
                        if (dataMep == null)
                            continue;

                        if (dataMep._Vertical01 != null && dataMep._Vertical01.IsValidObject == true)
                        {
                            Global.UIDoc.Document.Delete(dataMep._Vertical01.Id);
                            dataMep._Vertical01 = null;
                        }

                        if (dataMep._Vertical02 != null && dataMep._Vertical02.IsValidObject == true)
                        {
                            Global.UIDoc.Document.Delete(dataMep._Vertical02.Id);
                            dataMep._Vertical02 = null;
                        }

                        if (dataMep._Elbow01 != null && dataMep._Elbow01.IsValidObject == true)
                        {
                            Global.UIDoc.Document.Delete(dataMep._Elbow01.Id);
                            dataMep._Elbow01 = null;
                        }

                        if (dataMep._Elbow02 != null && dataMep._Elbow02.IsValidObject == true)
                        {
                            Global.UIDoc.Document.Delete(dataMep._Elbow02.Id);
                            dataMep._Elbow02 = null;
                        }

                        if (dataMep._Elbow03 != null && dataMep._Elbow03.IsValidObject == true)
                        {
                            Global.UIDoc.Document.Delete(dataMep._Elbow03.Id);
                            dataMep._Elbow03 = null;
                        }

                        if (dataMep._Elbow04 != null && dataMep._Elbow04.IsValidObject == true)
                        {
                            Global.UIDoc.Document.Delete(dataMep._Elbow04.Id);
                            dataMep._Elbow04 = null;
                        }
                    }

                    if (App.m_HolyUpDownForm.NotApply)
                    {
                        var lstSel = GetSelectedMEP();
                        foreach (MEPCurve mepCurve in lstSel)
                        {
                            var dataMep = _Datas.Find(item => item._MEPCurve.Id == mepCurve.Id);
                            if (dataMep == null)
                                continue;

                            double offset = dataMep._OldOffset / 304.8;

                            var loc = mepCurve.Location as LocationCurve;
                            if (loc == null)
                                continue;

                            var point = OffsetZ((loc.Curve as Line).Origin, offset);

                            var vectorMove = point - (loc.Curve as Line).Origin;

                            ElementTransformUtils.MoveElement(Global.UIDoc.Document, mepCurve.Id, vectorMove);
                        }
                    }
                    else
                    {
                        foreach (KeyValuePair<MEPCurve, List<FamilyInstance>> keyPair in data)
                        {
                            var mepCurve = keyPair.Key;

                            Level level = null;

                            var levelPara = mepCurve.LookupParameter("Reference Level");
                            if (levelPara != null)
                            {
                                level = Global.UIDoc.Document.GetElement(levelPara.AsElementId()) as Level;
                            }
                            var unions = keyPair.Value;

                            bool exist = false;

                            if (data_exist.ContainsKey(mepCurve))
                                exist = data_exist[mepCurve];

                            var dataMep = _Datas.Find(item => item._MEPCurve.Id == keyPair.Key.Id);
                            if (dataMep == null)
                                continue;

                            if (App.m_HolyUpDownForm.Elbow90)
                            {
                                if (exist)
                                {
                                    if (dataMep._UnionPoint01 != null)
                                    {
                                        if (dataMep._Vertical01 != null && dataMep._Vertical01.IsValidObject)
                                            Global.UIDoc.Document.Delete(dataMep._Vertical01.Id);

                                        var p2 = OffsetZ(dataMep._UnionPoint01, dataMep._OldOffset);
                                        dataMep._Vertical01 = MEPUtilscs.CC(mepCurve, Line.CreateBound(dataMep._UnionPoint01, p2));

                                        if (dataMep._Elbow01 != null && dataMep._Elbow01.IsValidObject)
                                            Global.UIDoc.Document.Delete(dataMep._Elbow01.Id);

                                        dataMep._Elbow01 = MEPUtilscs.CE(dataMep._Other01, dataMep._Vertical01, dataMep._UnionPoint01);

                                        if (dataMep._Elbow02 != null && dataMep._Elbow02.IsValidObject)
                                            Global.UIDoc.Document.Delete(dataMep._Elbow02.Id);
                                        dataMep._Elbow02 = MEPUtilscs.CE(mepCurve, dataMep._Vertical01, p2);
                                    }
                                    if (dataMep._UnionPoint02 != null)
                                    {
                                        if (dataMep._Vertical02 != null && dataMep._Vertical02.IsValidObject)
                                            Global.UIDoc.Document.Delete(dataMep._Vertical02.Id);

                                        var p2 = OffsetZ(dataMep._UnionPoint02, dataMep._OldOffset);
                                        dataMep._Vertical02 = MEPUtilscs.CC(mepCurve, Line.CreateBound(dataMep._UnionPoint02, p2));

                                        if (dataMep._Elbow03 != null && dataMep._Elbow03.IsValidObject)
                                            Global.UIDoc.Document.Delete(dataMep._Elbow03.Id);

                                        dataMep._Elbow03 = MEPUtilscs.CE(dataMep._Other02, dataMep._Vertical02, dataMep._UnionPoint02);

                                        if (dataMep._Elbow04 != null && dataMep._Elbow04.IsValidObject)
                                            Global.UIDoc.Document.Delete(dataMep._Elbow04.Id);
                                        dataMep._Elbow04 = MEPUtilscs.CE(mepCurve, dataMep._Vertical02, p2);
                                    }
                                }
                                else
                                {
                                    int index = 0;
                                    foreach (FamilyInstance union in unions)
                                    {
                                        var locationPoint = (union.Location as LocationPoint).Point;

                                        if (index == 0)
                                        {
                                            dataMep._UnionPoint01 = locationPoint;
                                        }
                                        else
                                        {
                                            dataMep._UnionPoint02 = locationPoint;
                                        }

                                        var p2 = OffsetZ(locationPoint, dataMep._OldOffset);

                                        if (dataMep._Vertical01 == null && index == 0)
                                        {
                                            dataMep._Vertical01 = MEPUtilscs.CC(mepCurve, Line.CreateBound(locationPoint, p2));
                                        }
                                        else
                                        {
                                            dataMep._Vertical02 = MEPUtilscs.CC(mepCurve, Line.CreateBound(locationPoint, p2));
                                        }

                                        MEPCurve mepOther = null;
                                        foreach (Connector c in union.MEPModel.ConnectorManager.Connectors)
                                        {
                                            foreach (Connector cmep in c.AllRefs)
                                            {
                                                if (cmep.Owner.Id == mepCurve.Id)
                                                    continue;

                                                if (cmep.Owner as MEPCurve != null)
                                                {
                                                    mepOther = cmep.Owner as MEPCurve;
                                                    break;
                                                }
                                            }
                                        }
                                        if (mepOther != null)
                                        {
                                            if (index == 0)
                                            {
                                                var elbow1 = MEPUtilscs.CE(mepOther, dataMep._Vertical01, locationPoint);
                                                dataMep._Elbow01 = (elbow1 as FamilyInstance);

                                                var elbow2 = MEPUtilscs.CE(mepCurve, dataMep._Vertical01, p2);
                                                dataMep._Elbow02 = (elbow2 as FamilyInstance);

                                                dataMep._Other01 = mepOther;
                                            }
                                            else
                                            {
                                                var elbow3 = MEPUtilscs.CE(mepOther, dataMep._Vertical02, locationPoint);
                                                dataMep._Elbow03 = (elbow3 as FamilyInstance);

                                                var elbow4 = MEPUtilscs.CE(mepCurve, dataMep._Vertical02, p2);
                                                dataMep._Elbow04 = (elbow4 as FamilyInstance);

                                                dataMep._Other02 = mepOther;
                                            }
                                        }
                                        Global.UIDoc.Document.Delete(union.Id);

                                        index++;
                                    }
                                }
                            }
                            else
                            {
                                double split = dataMep._OldOffset;
                                if (App.m_HolyUpDownForm.ElbowCustom)
                                {
                                    var cgv = dataMep._OldOffset;
                                    var radian = MyUnit.da(App.m_HolyUpDownForm.AngleCustom);
                                    var ch = cgv / Math.Sin(radian);
                                    split = ch * Math.Cos(radian);
                                }

                                if (exist == true)
                                {
                                    var curve = (mepCurve.Location as LocationCurve).Curve;

                                    if (dataMep._UnionPoint01 != null)
                                    {
                                        var p0 = curve.GetEndPoint(0);
                                        var p1 = curve.GetEndPoint(1);

                                        var line = Line.CreateBound(p0, p1);
                                        if (line.Length <= Math.Abs(split))
                                        {
                                            MessageBox.Show("Offset is invalid.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                            return Result.Cancelled;
                                        }

                                        if (dataMep._Vertical01 != null && dataMep._Vertical01.IsValidObject)
                                            Global.UIDoc.Document.Delete(dataMep._Vertical01.Id);

                                        if (dataMep._Elbow01 != null && dataMep._Elbow01.IsValidObject)
                                            Global.UIDoc.Document.Delete(dataMep._Elbow01.Id);

                                        if (dataMep._Elbow02 != null && dataMep._Elbow02.IsValidObject)
                                            Global.UIDoc.Document.Delete(dataMep._Elbow02.Id);

                                        var locationPoint = dataMep._UnionPoint01;

                                        XYZ newPoint = null;
                                        var p3 = OffsetZ(locationPoint, dataMep._OldOffset);

                                        if (p0.DistanceTo(p3) < p1.DistanceTo(p3))
                                        {
                                            newPoint = line.Evaluate(Math.Abs(split), false);
                                            (mepCurve.Location as LocationCurve).Curve = Line.CreateBound(newPoint, p1);
                                        }
                                        else
                                        {
                                            newPoint = line.Evaluate(line.Length - Math.Abs(split), false);
                                            (mepCurve.Location as LocationCurve).Curve = Line.CreateBound(p0, newPoint);
                                        }

                                        dataMep._Vertical01 = MEPUtilscs.CC(mepCurve, Line.CreateBound(locationPoint, newPoint));

                                        if (dataMep._Other01 != null)
                                        {
                                            dataMep._Elbow01 = MEPUtilscs.CE(dataMep._Other01, dataMep._Vertical01, locationPoint);
                                            dataMep._Elbow02 = MEPUtilscs.CE(mepCurve, dataMep._Vertical01, newPoint);
                                        }
                                    }

                                    if (dataMep._UnionPoint02 != null)
                                    {
                                        //Get again location
                                        curve = (mepCurve.Location as LocationCurve).Curve;
                                        var p0 = curve.GetEndPoint(0);
                                        var p1 = curve.GetEndPoint(1);

                                        var line = Line.CreateBound(p0, p1);
                                        if (line.Length <= Math.Abs(split))
                                        {
                                            MessageBox.Show("Offset is invalid.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                            return Result.Cancelled;
                                        }

                                        if (dataMep._Vertical02 != null && dataMep._Vertical02.IsValidObject)
                                            Global.UIDoc.Document.Delete(dataMep._Vertical02.Id);

                                        if (dataMep._Elbow03 != null && dataMep._Elbow03.IsValidObject)
                                            Global.UIDoc.Document.Delete(dataMep._Elbow03.Id);

                                        if (dataMep._Elbow04 != null && dataMep._Elbow04.IsValidObject)
                                            Global.UIDoc.Document.Delete(dataMep._Elbow04.Id);

                                        var locationPoint = dataMep._UnionPoint02;

                                        XYZ newPoint = null;
                                        var p3 = OffsetZ(locationPoint, dataMep._OldOffset);

                                        if (p0.DistanceTo(p3) < p1.DistanceTo(p3))
                                        {
                                            newPoint = line.Evaluate(Math.Abs(split), false);
                                            (mepCurve.Location as LocationCurve).Curve = Line.CreateBound(newPoint, p1);
                                        }
                                        else
                                        {
                                            newPoint = line.Evaluate(line.Length - Math.Abs(split), false);
                                            (mepCurve.Location as LocationCurve).Curve = Line.CreateBound(p0, newPoint);
                                        }

                                        dataMep._Vertical02 = MEPUtilscs.CC(mepCurve, Line.CreateBound(locationPoint, newPoint));

                                        if (dataMep._Other02 != null)
                                        {
                                            dataMep._Elbow03 = MEPUtilscs.CE(dataMep._Other02, dataMep._Vertical02, locationPoint);
                                            dataMep._Elbow04 = MEPUtilscs.CE(mepCurve, dataMep._Vertical02, newPoint);
                                        }
                                    }
                                }
                                else
                                {
                                    int index = 0;
                                    foreach (FamilyInstance union in unions)
                                    {
                                        var curve = (mepCurve.Location as LocationCurve).Curve;
                                        var p0 = curve.GetEndPoint(0);
                                        var p1 = curve.GetEndPoint(1);

                                        var line = Line.CreateBound(p0, p1);
                                        if (line.Length <= Math.Abs(split))
                                        {
                                            MessageBox.Show("Offset is invalid.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                            return Result.Cancelled;
                                        }

                                        var locationPoint = (union.Location as LocationPoint).Point;

                                        if (index == 0)
                                        {
                                            dataMep._UnionPoint01 = locationPoint;
                                        }
                                        else
                                        {
                                            dataMep._UnionPoint02 = locationPoint;
                                        }

                                        XYZ newPoint = null;

                                        var p3 = OffsetZ(locationPoint, dataMep._OldOffset);

                                        Line newCurve = null;
                                        if (p0.DistanceTo(p3) < p1.DistanceTo(p3))
                                        {
                                            newPoint = line.Evaluate(Math.Abs(split), false);
                                            newCurve = Line.CreateBound(newPoint, p1);
                                        }
                                        else
                                        {
                                            newPoint = line.Evaluate(line.Length - Math.Abs(split), false);
                                            newCurve = Line.CreateBound(p0, newPoint);
                                        }

                                        (mepCurve.Location as LocationCurve).Curve = newCurve;

                                        //Must set Middle Elevation
                                        var midleEle = mepCurve.LookupParameter("Middle Elevation");
                                        if (midleEle != null)
                                        {
                                            var mid = (newCurve.GetEndPoint(0).Z + newCurve.GetEndPoint(1).Z) / 2;

                                            if (level != null)
                                            {
                                                mid -= level.Elevation;
                                            }

                                            midleEle.Set(mid);
                                        }

                                        if (index == 0)
                                            dataMep._Vertical01 = MEPUtilscs.CC(mepCurve, Line.CreateBound(locationPoint, newPoint));
                                        else
                                            dataMep._Vertical02 = MEPUtilscs.CC(mepCurve, Line.CreateBound(locationPoint, newPoint));

                                        MEPCurve mepOther = null;
                                        foreach (Connector c in union.MEPModel.ConnectorManager.Connectors)
                                        {
                                            foreach (Connector cmep in c.AllRefs)
                                            {
                                                if (cmep.Owner.Id == mepCurve.Id)
                                                    continue;

                                                if (cmep.Owner as MEPCurve != null)
                                                {
                                                    mepOther = cmep.Owner as MEPCurve;
                                                    break;
                                                }
                                            }
                                        }
                                        if (mepOther != null)
                                        {
                                            if (index == 0)
                                            {
                                                var elbow1 = MEPUtilscs.CE(mepOther, dataMep._Vertical01, locationPoint);
                                                dataMep._Elbow01 = (elbow1 as FamilyInstance);

                                                var elbow2 = MEPUtilscs.CE(mepCurve, dataMep._Vertical01, newPoint);
                                                dataMep._Elbow02 = (elbow2 as FamilyInstance);

                                                dataMep._Other01 = mepOther;
                                            }
                                            else
                                            {
                                                var elbow3 = MEPUtilscs.CE(mepOther, dataMep._Vertical02, locationPoint);
                                                dataMep._Elbow03 = (elbow3 as FamilyInstance);

                                                var elbow4 = MEPUtilscs.CE(mepCurve, dataMep._Vertical02, newPoint);
                                                dataMep._Elbow04 = (elbow4 as FamilyInstance);

                                                dataMep._Other02 = mepOther;
                                            }
                                        }

                                        Global.UIDoc.Document.Delete(union.Id);
                                        index++;
                                    }
                                }
                            }
                        }
                    }
                }

                //Add to storage

                var values = WriteXML();
                bool result = StorageUtility.AddEntity(Global.UIDoc.Document.ProjectInformation, StorageUtility.m_MEP_HoLyUpDown_Guild, StorageUtility.m_MEP_HoLyUpDown, values);

                tran.Commit();

                return Result.Succeeded;
            }
            catch (System.Exception ex)
            {
                return Result.Cancelled;
            }
        }

        public static Result Run_OKRefreshDara()
        {
            try
            {
                using (Transaction reTrans = new Transaction(Global.UIDoc.Document, "DELETE_HOLY_UPDOWN_SCHEMA"))
                {
                    reTrans.Start();
                    try
                    {
                        string saveValue = string.Empty;
                        StorageUtility.SetValue(Global.UIDoc.Document.ProjectInformation, StorageUtility.m_MEP_HoLyUpDown_Guild, StorageUtility.m_MEP_HoLyUpDown, typeof(string), saveValue);
                        reTrans.Commit();
                    }
                    catch (Exception)
                    {
                        reTrans.RollBack();
                    }
                }
            }
            catch { }
            return Result.Cancelled;
        }

        private static XYZ OffsetZ(XYZ point, double offset)
        {
            return new XYZ(point.X, point.Y, point.Z + offset);
        }

        private static List<MEPCurve> GetSelectedMEP()
        {
            try
            {
                List<MEPCurve> meps = new List<MEPCurve>();
                var ids = Global.UIDoc.Selection.GetElementIds();
                foreach (ElementId id in ids)
                {
                    var element = Global.UIDoc.Document.GetElement(id) as MEPCurve;
                    if (element != null)
                        meps.Add(element);
                }
                return meps;
            }
            catch (System.Exception ex)
            {
                return null;
            }
        }

        private static List<FamilyInstance> GetSelectedElbow()
        {
            try
            {
                List<FamilyInstance> lstElbow = new List<FamilyInstance>();
                var ids = Global.UIDoc.Selection.GetElementIds();
                foreach (ElementId id in ids)
                {
                    var element = Global.UIDoc.Document.GetElement(id) as FamilyInstance;
                    if (element != null)
                        lstElbow.Add(element);
                }
                return lstElbow;
            }
            catch (System.Exception ex)
            {
                return null;
            }
        }

        public static void ReadData(string xml)
        {
            try
            {
                _Datas.Clear();

                if (!string.IsNullOrWhiteSpace(xml))
                {
                    XmlDocument xmlDoc = new XmlDocument();
                    xmlDoc.LoadXml(xml);

                    var dataNodes = xmlDoc.SelectNodes("/Datas/Data");

                    foreach (XmlNode node in dataNodes)
                    {
                        XmlAttribute attribute = node.Attributes["MEPCurve"];
                        if (attribute == null || attribute.Value == null)
                            continue;

                        var mepCurve = Global.UIDoc.Document.GetElement(attribute.Value) as MEPCurve;
                        if (mepCurve == null)
                            continue;

                        Data dataMep = new Data(mepCurve);

                        attribute = node.Attributes["Start"];
                        if (attribute == null || attribute.Value == null)
                            continue;
                        dataMep._Start = Read(attribute.Value);

                        attribute = node.Attributes["End"];
                        if (attribute == null || attribute.Value == null)
                            continue;
                        dataMep._End = Read(attribute.Value);

                        attribute = node.Attributes["PStart"];
                        if (attribute == null || attribute.Value == null)
                            continue;
                        dataMep._PStart = Read(attribute.Value);

                        attribute = node.Attributes["PEnd"];
                        if (attribute == null || attribute.Value == null)
                            continue;
                        dataMep._PEnd = Read(attribute.Value);

                        attribute = node.Attributes["UnionPoint01"];
                        dataMep._UnionPoint01 = Read(attribute.Value);

                        attribute = node.Attributes["UnionPoint02"];
                        dataMep._UnionPoint02 = Read(attribute.Value);

                        attribute = node.Attributes["Elbow01"];
                        if (attribute != null && attribute.Value != null)
                        {
                            dataMep._Elbow01 = Global.UIDoc.Document.GetElement(attribute.Value) as Element;
                        }
                        attribute = node.Attributes["Elbow02"];
                        if (attribute != null && attribute.Value != null)
                        {
                            dataMep._Elbow02 = Global.UIDoc.Document.GetElement(attribute.Value) as Element;
                        }

                        attribute = node.Attributes["Elbow03"];
                        if (attribute != null && attribute.Value != null)
                        {
                            dataMep._Elbow03 = Global.UIDoc.Document.GetElement(attribute.Value) as Element;
                        }
                        attribute = node.Attributes["Elbow04"];
                        if (attribute != null && attribute.Value != null)
                        {
                            dataMep._Elbow04 = Global.UIDoc.Document.GetElement(attribute.Value) as Element;
                        }

                        attribute = node.Attributes["Vertical01"];
                        if (attribute != null && attribute.Value != null)
                        {
                            dataMep._Vertical01 = Global.UIDoc.Document.GetElement(attribute.Value) as MEPCurve;
                        }

                        attribute = node.Attributes["Vertical02"];
                        if (attribute != null && attribute.Value != null)
                        {
                            dataMep._Vertical02 = Global.UIDoc.Document.GetElement(attribute.Value) as MEPCurve;
                        }

                        attribute = node.Attributes["Other02"];
                        if (attribute != null && attribute.Value != null)
                        {
                            dataMep._Other02 = Global.UIDoc.Document.GetElement(attribute.Value) as MEPCurve;
                        }

                        attribute = node.Attributes["Other01"];
                        if (attribute != null && attribute.Value != null)
                        {
                            dataMep._Other01 = Global.UIDoc.Document.GetElement(attribute.Value) as MEPCurve;
                        }

                        attribute = node.Attributes["Other02"];
                        if (attribute != null && attribute.Value != null)
                        {
                            dataMep._Other02 = Global.UIDoc.Document.GetElement(attribute.Value) as MEPCurve;
                        }

                        attribute = node.Attributes["OldOffset"];
                        if (attribute != null && attribute.Value != null)
                        {
                            dataMep._OldOffset = ToDouble(attribute.Value);
                        }

                        _Datas.Add(dataMep);
                    }
                }
            }
            catch (Exception)
            { }
        }

        private static XYZ Read(string stringValue)
        {
            stringValue = stringValue.Replace("(", "");
            stringValue = stringValue.Replace(")", "");

            var splits = stringValue.Split(',');

            if (splits.Length != 3)
                return null;

            var x = splits[0].Trim();
            var y = splits[1].Trim();
            var z = splits[2].Trim();

            double dx = Convert.ToDouble(x);
            double dy = Convert.ToDouble(y);
            double dz = Convert.ToDouble(z);

            return new XYZ(dx, dy, dz);
        }

        public static double ToDouble(string input)
        {
            double output = 0.0;
            if (double.TryParse(input, out double space))
            {
                output = space;
            }
            return output;
        }

        public static string WriteXML()
        {
            XmlDocument xmlDoc = new XmlDocument();
            XmlNode rootNode = xmlDoc.CreateElement("Datas");
            xmlDoc.AppendChild(rootNode);

            foreach (Data data in _Datas)
            {
                if (data._MEPCurve == null || data._Start == null || data._End == null)
                    continue;

                XmlNode dataNode = xmlDoc.CreateElement("Data");
                XmlAttribute attribute = xmlDoc.CreateAttribute("MEPCurve");
                attribute.Value = data._MEPCurve.UniqueId.ToString();
                dataNode.Attributes.Append(attribute);

                attribute = xmlDoc.CreateAttribute("Start");
                if (data._Start != null)
                    attribute.Value = data._Start.ToString();
                else
                    attribute.Value = string.Empty;

                dataNode.Attributes.Append(attribute);

                attribute = xmlDoc.CreateAttribute("End");
                if (data._End != null)
                    attribute.Value = data._End.ToString();
                else
                    attribute.Value = string.Empty;
                dataNode.Attributes.Append(attribute);

                attribute = xmlDoc.CreateAttribute("PStart");
                if (data._PStart != null)
                    attribute.Value = data._PStart.ToString();
                else
                    attribute.Value = string.Empty;
                dataNode.Attributes.Append(attribute);

                attribute = xmlDoc.CreateAttribute("PEnd");
                if (data._PEnd != null)
                    attribute.Value = data._PEnd.ToString();
                else
                    attribute.Value = string.Empty;
                dataNode.Attributes.Append(attribute);

                attribute = xmlDoc.CreateAttribute("UnionPoint01");
                if (data._UnionPoint01 != null)
                    attribute.Value = data._UnionPoint01.ToString();
                else
                    attribute.Value = string.Empty;
                dataNode.Attributes.Append(attribute);

                attribute = xmlDoc.CreateAttribute("UnionPoint02");
                if (data._UnionPoint02 != null)
                    attribute.Value = data._UnionPoint02.ToString();
                else
                    attribute.Value = string.Empty;
                dataNode.Attributes.Append(attribute);

                attribute = xmlDoc.CreateAttribute("Elbow01");
                if (data._Elbow01 != null)
                    attribute.Value = data._Elbow01.UniqueId.ToString();
                else
                    attribute.Value = string.Empty;
                dataNode.Attributes.Append(attribute);

                attribute = xmlDoc.CreateAttribute("Elbow02");
                if (data._Elbow02 != null)
                    attribute.Value = data._Elbow02.UniqueId.ToString();
                else
                    attribute.Value = string.Empty;
                dataNode.Attributes.Append(attribute);

                attribute = xmlDoc.CreateAttribute("Elbow03");
                if (data._Elbow03 != null)
                    attribute.Value = data._Elbow03.UniqueId.ToString();
                else
                    attribute.Value = string.Empty;
                dataNode.Attributes.Append(attribute);

                attribute = xmlDoc.CreateAttribute("Elbow04");
                if (data._Elbow04 != null)
                    attribute.Value = data._Elbow04.UniqueId.ToString();
                else
                    attribute.Value = string.Empty;
                dataNode.Attributes.Append(attribute);

                attribute = xmlDoc.CreateAttribute("Vertical01");
                if (data._Vertical01 != null)
                    attribute.Value = data._Vertical01.UniqueId.ToString();
                else
                    attribute.Value = string.Empty;
                dataNode.Attributes.Append(attribute);

                attribute = xmlDoc.CreateAttribute("Vertical02");
                if (data._Vertical02 != null)
                    attribute.Value = data._Vertical02.UniqueId.ToString();
                else
                    attribute.Value = string.Empty;
                dataNode.Attributes.Append(attribute);

                attribute = xmlDoc.CreateAttribute("Other01");
                if (data._Other01 != null)
                    attribute.Value = data._Other01.UniqueId.ToString();
                else
                    attribute.Value = string.Empty;
                dataNode.Attributes.Append(attribute);

                attribute = xmlDoc.CreateAttribute("Other02");
                if (data._Other02 != null)
                    attribute.Value = data._Other02.UniqueId.ToString();
                else
                    attribute.Value = string.Empty;
                dataNode.Attributes.Append(attribute);

                attribute = xmlDoc.CreateAttribute("OldOffset");
                if (data._OldOffset != double.NaN)
                    attribute.Value = data._OldOffset.ToString();
                else
                    attribute.Value = string.Empty;
                dataNode.Attributes.Append(attribute);

                rootNode.AppendChild(dataNode);
            }

            using (var stringWriter = new StringWriter())
            using (var xmlTextWriter = XmlWriter.Create(stringWriter))
            {
                xmlDoc.WriteTo(xmlTextWriter);
                xmlTextWriter.Flush();
                return stringWriter.GetStringBuilder().ToString();
            }
        }

        public static Connector GetConnectorClosestTo(
         Element e,
         XYZ p)
        {
            ConnectorManager cm = GetConnectorManager(e);

            return null == cm
              ? null
              : GetConnectorClosestTo(cm.Connectors, p);
        }

        private static ConnectorManager GetConnectorManager(
        Element e)
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

        private static Connector GetConnectorClosestTo(
        ConnectorSet connectors,
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

        public static Result Run_UpDownStep(bool downStep = false)
        {
            try
            {
                if (App.m_HolyUpDownForm == null)
                    return Result.Cancelled;

                List<MEPCurve> mepCurveSelects = GetSelectedMEP();
                if (mepCurveSelects == null || mepCurveSelects.Count == 0)
                    return Result.Cancelled;

                double dStepValue = UnitUtils.ConvertToInternalUnits(Math.Abs(App.m_HolyUpDownForm.UpStepValue), DisplayUnitType.DUT_MILLIMETERS);

                foreach (MEPCurve mepCurve in mepCurveSelects)
                {
                    double dOldOffsetParamVal = GetBuiltInParameterValue(mepCurve, BuiltInParameter.RBS_OFFSET_PARAM)
                                                != null ? (double)GetBuiltInParameterValue(mepCurve, BuiltInParameter.RBS_OFFSET_PARAM) : double.MinValue;

                    if (dOldOffsetParamVal == double.MinValue)
                        continue;
                    using (Transaction reTrans = new Transaction(Global.UIDoc.Document, "HOLYUPDOWN_UPSTEP"))
                    {
                        if (reTrans.Start() == TransactionStatus.Started)
                        {
                            try
                            {
                                double dNewOffsetParamVal = downStep == false ? dOldOffsetParamVal + dStepValue : dOldOffsetParamVal - dStepValue;
                                SetBuiltinParameterValue(mepCurve, BuiltInParameter.RBS_OFFSET_PARAM, dNewOffsetParamVal);
                                reTrans.Commit();
                            }
                            catch (Exception)
                            {
                                reTrans.RollBack();
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
            }
            return Result.Cancelled;
        }

        public static Result Run_UpDownElbowControl(bool downElbowControl = false)
        {
            using (Transaction tran = new Transaction(Global.UIDoc.Document, "Elbow Control"))
            {
                tran.Start();
                try
                {
                    if (App.m_HolyUpDownForm == null)
                        return Result.Cancelled;

                    var lstElbow = GetSelectedElbow();
                    if (lstElbow == null || lstElbow.Count == 0)
                        return Result.Cancelled;

                    double offset = App.m_HolyUpDownForm.UpElbowStepValue / 304.8;

                    Dictionary<FamilyInstance, XYZ> keyValuePairs = new Dictionary<FamilyInstance, XYZ>();
                    foreach (var elbow in lstElbow)
                    {
                        var lcPoint = elbow.Location as LocationPoint;
                        if (lcPoint == null)
                            continue;

                        keyValuePairs.Add(elbow, lcPoint.Point);
                    }

                    List<FamilyInstance> lstElbowY = new List<FamilyInstance>();
                    List<FamilyInstance> lstElbowX = new List<FamilyInstance>();
                    List<FamilyInstance> lstElbowC = new List<FamilyInstance>();
                    foreach (var item in keyValuePairs)
                    {
                        foreach (var item1 in keyValuePairs)
                        {
                            if (item.Key.Id == item1.Key.Id)
                                continue;

                            if (Common.Equals(item.Value.Y, item1.Value.Y))
                            {
                                if (!lstElbowY.Select(x => x.Id).Contains(item.Key.Id))
                                    lstElbowY.Add(item.Key);
                                if (!lstElbowY.Select(x => x.Id).Contains(item1.Key.Id))
                                    lstElbowY.Add(item1.Key);
                            }
                            else if (Common.Equals(item.Value.X, item1.Value.X))
                            {
                                if (!lstElbowX.Select(x => x.Id).Contains(item.Key.Id))
                                    lstElbowX.Add(item.Key);
                                if (!lstElbowX.Select(x => x.Id).Contains(item1.Key.Id))
                                    lstElbowX.Add(item1.Key);
                            }
                        }
                    }

                    lstElbowY = lstElbowY.OrderByDescending(x => (x.Location as LocationPoint).Point.Z).ToList();
                    lstElbowX = lstElbowX.OrderByDescending(x => (x.Location as LocationPoint).Point.Z).ToList();

                    var d = keyValuePairs.Where(x => !lstElbowX.Select(y => y.Id).Contains(x.Key.Id) && !lstElbowY.Select(y => y.Id).Contains(x.Key.Id));

                    foreach (var item in d)
                    {
                        lstElbowC.Add(item.Key);
                    }

                    foreach (var item in lstElbowC)
                    {
                        var lstCon = GetListConnector(item.MEPModel.ConnectorManager.Connectors);
                        var vector = lstCon[1].Origin - lstCon[0].Origin;
                        if (vector.X == 0)
                        {
                            if (downElbowControl)
                            {
                                XYZ a = new XYZ((item.Location as LocationPoint).Point.X, (item.Location as LocationPoint).Point.Y - offset, (item.Location as LocationPoint).Point.Z);

                                ElementTransformUtils.MoveElement(Global.UIDoc.Document, item.Id, a - (item.Location as LocationPoint).Point);
                            }
                            else
                            {
                                XYZ a = new XYZ((item.Location as LocationPoint).Point.X, (item.Location as LocationPoint).Point.Y + offset, (item.Location as LocationPoint).Point.Z);

                                ElementTransformUtils.MoveElement(Global.UIDoc.Document, item.Id, a - (item.Location as LocationPoint).Point);
                            }
                        }
                        else if (vector.Y == 0)
                        {
                            if (!downElbowControl)
                            {
                                XYZ a = new XYZ((item.Location as LocationPoint).Point.X + offset, (item.Location as LocationPoint).Point.Y, (item.Location as LocationPoint).Point.Z);

                                ElementTransformUtils.MoveElement(Global.UIDoc.Document, item.Id, a - (item.Location as LocationPoint).Point);
                            }
                            else
                            {
                                XYZ a = new XYZ((item.Location as LocationPoint).Point.X - offset, (item.Location as LocationPoint).Point.Y, (item.Location as LocationPoint).Point.Z);

                                ElementTransformUtils.MoveElement(Global.UIDoc.Document, item.Id, a - (item.Location as LocationPoint).Point);
                            }
                        }
                    }

                    var dicX = GetDicElbow(lstElbowX);
                    var dicY = GetDicElbow(lstElbowY);

                    MoveElbowY(Global.UIDoc.Document, dicY, offset, downElbowControl);

                    MoveElbowX(Global.UIDoc.Document, dicX, offset, downElbowControl);

                    tran.Commit();
                }
                catch (Exception)
                {
                    tran.RollBack();
                }
            }
            return Result.Cancelled;
        }

        private static void MoveElbowY(Document doc, Dictionary<int, List<FamilyInstance>> dicY, double offset, bool isDown)
        {
            foreach (var item in dicY)
            {
                if (item.Value.Count > 1)
                {
                    for (int i = 0; i < item.Value.Count; i++)
                    {
                        if (i == item.Value.Count - 1)
                        {
                            var vector = (item.Value[0].Location as LocationPoint).Point - (item.Value[1].Location as LocationPoint).Point;
                            if (vector.X < 0)
                            {
                                if (isDown)
                                {
                                    XYZ a = new XYZ((item.Value[i].Location as LocationPoint).Point.X - offset, (item.Value[i].Location as LocationPoint).Point.Y, (item.Value[i].Location as LocationPoint).Point.Z);
                                    ElementTransformUtils.MoveElement(doc, item.Value[i].Id, a - (item.Value[i].Location as LocationPoint).Point);
                                }
                                else
                                {
                                    XYZ a = new XYZ((item.Value[i].Location as LocationPoint).Point.X + offset, (item.Value[i].Location as LocationPoint).Point.Y, (item.Value[i].Location as LocationPoint).Point.Z);
                                    ElementTransformUtils.MoveElement(doc, item.Value[i].Id, a - (item.Value[i].Location as LocationPoint).Point);
                                }
                            }
                            else
                            {
                                if (isDown)
                                {
                                    XYZ a = new XYZ((item.Value[i].Location as LocationPoint).Point.X + offset, (item.Value[i].Location as LocationPoint).Point.Y, (item.Value[i].Location as LocationPoint).Point.Z);

                                    ElementTransformUtils.MoveElement(doc, item.Value[i].Id, a - (item.Value[i].Location as LocationPoint).Point);
                                }
                                else
                                {
                                    XYZ a = new XYZ((item.Value[i].Location as LocationPoint).Point.X - offset, (item.Value[i].Location as LocationPoint).Point.Y, (item.Value[i].Location as LocationPoint).Point.Z);

                                    ElementTransformUtils.MoveElement(doc, item.Value[i].Id, a - (item.Value[i].Location as LocationPoint).Point);
                                }
                            }
                        }
                        else
                        {
                            var vector = (item.Value[1].Location as LocationPoint).Point - (item.Value[0].Location as LocationPoint).Point;
                            if (vector.X > 0)
                            {
                                if (isDown)
                                {
                                    XYZ a = new XYZ((item.Value[i].Location as LocationPoint).Point.X + offset, (item.Value[i].Location as LocationPoint).Point.Y, (item.Value[i].Location as LocationPoint).Point.Z);

                                    ElementTransformUtils.MoveElement(doc, item.Value[i].Id, a - (item.Value[i].Location as LocationPoint).Point);
                                }
                                else
                                {
                                    XYZ a = new XYZ((item.Value[i].Location as LocationPoint).Point.X - offset, (item.Value[i].Location as LocationPoint).Point.Y, (item.Value[i].Location as LocationPoint).Point.Z);

                                    ElementTransformUtils.MoveElement(doc, item.Value[i].Id, a - (item.Value[i].Location as LocationPoint).Point);
                                }
                            }
                            else
                            {
                                if (isDown)
                                {
                                    XYZ a = new XYZ((item.Value[i].Location as LocationPoint).Point.X - offset, (item.Value[i].Location as LocationPoint).Point.Y, (item.Value[i].Location as LocationPoint).Point.Z);

                                    ElementTransformUtils.MoveElement(doc, item.Value[i].Id, a - (item.Value[i].Location as LocationPoint).Point);
                                }
                                else
                                {
                                    XYZ a = new XYZ((item.Value[i].Location as LocationPoint).Point.X + offset, (item.Value[i].Location as LocationPoint).Point.Y, (item.Value[i].Location as LocationPoint).Point.Z);

                                    ElementTransformUtils.MoveElement(doc, item.Value[i].Id, a - (item.Value[i].Location as LocationPoint).Point);
                                }
                            }
                        }
                    }
                }
                else
                {
                    List<Connector> lstCon = GetListConnector(item.Value.FirstOrDefault().MEPModel.ConnectorManager.Connectors);
                    if (lstCon?.Count > 0)
                    {
                        var vector = lstCon[1].Origin - lstCon[0].Origin;
                        if (vector.X > 0)
                        {
                            if (isDown)
                            {
                                XYZ a = new XYZ((item.Value.FirstOrDefault().Location as LocationPoint).Point.X + offset, (item.Value.FirstOrDefault().Location as LocationPoint).Point.Y, (item.Value.FirstOrDefault().Location as LocationPoint).Point.Z);

                                ElementTransformUtils.MoveElement(doc, item.Value.FirstOrDefault().Id, a - (item.Value.FirstOrDefault().Location as LocationPoint).Point);
                            }
                            else
                            {
                                XYZ a = new XYZ((item.Value.FirstOrDefault().Location as LocationPoint).Point.X - offset, (item.Value.FirstOrDefault().Location as LocationPoint).Point.Y, (item.Value.FirstOrDefault().Location as LocationPoint).Point.Z);

                                ElementTransformUtils.MoveElement(doc, item.Value.FirstOrDefault().Id, a - (item.Value.FirstOrDefault().Location as LocationPoint).Point);
                            }
                        }
                        else
                        {
                            if (isDown)
                            {
                                XYZ a = new XYZ((item.Value.FirstOrDefault().Location as LocationPoint).Point.X - offset, (item.Value.FirstOrDefault().Location as LocationPoint).Point.Y, (item.Value.FirstOrDefault().Location as LocationPoint).Point.Z);

                                ElementTransformUtils.MoveElement(doc, item.Value.FirstOrDefault().Id, a - (item.Value.FirstOrDefault().Location as LocationPoint).Point);
                            }
                            else
                            {
                                XYZ a = new XYZ((item.Value.FirstOrDefault().Location as LocationPoint).Point.X + offset, (item.Value.FirstOrDefault().Location as LocationPoint).Point.Y, (item.Value.FirstOrDefault().Location as LocationPoint).Point.Z);

                                ElementTransformUtils.MoveElement(doc, item.Value.FirstOrDefault().Id, a - (item.Value.FirstOrDefault().Location as LocationPoint).Point);
                            }
                        }
                    }
                }
            }
        }

        private static void MoveElbowX(Document doc, Dictionary<int, List<FamilyInstance>> dicX, double offset, bool isDown)
        {
            foreach (var item in dicX)
            {
                if (item.Value.Count > 1)
                {
                    for (int i = 0; i < item.Value.Count; i++)
                    {
                        if (i == item.Value.Count - 1)
                        {
                            var vector = (item.Value[0].Location as LocationPoint).Point - (item.Value[1].Location as LocationPoint).Point;
                            if (vector.Y < 0)
                            {
                                if (isDown)
                                {
                                    XYZ a = new XYZ((item.Value[i].Location as LocationPoint).Point.X, (item.Value[i].Location as LocationPoint).Point.Y - offset, (item.Value[i].Location as LocationPoint).Point.Z);

                                    ElementTransformUtils.MoveElement(doc, item.Value[i].Id, a - (item.Value[i].Location as LocationPoint).Point);
                                }
                                else
                                {
                                    XYZ a = new XYZ((item.Value[i].Location as LocationPoint).Point.X, (item.Value[i].Location as LocationPoint).Point.Y + offset, (item.Value[i].Location as LocationPoint).Point.Z);

                                    ElementTransformUtils.MoveElement(doc, item.Value[i].Id, a - (item.Value[i].Location as LocationPoint).Point);
                                }
                            }
                            else
                            {
                                if (isDown)
                                {
                                    XYZ a = new XYZ((item.Value[i].Location as LocationPoint).Point.X, (item.Value[i].Location as LocationPoint).Point.Y + offset, (item.Value[i].Location as LocationPoint).Point.Z);

                                    ElementTransformUtils.MoveElement(doc, item.Value[i].Id, a - (item.Value[i].Location as LocationPoint).Point);
                                }
                                else
                                {
                                    XYZ a = new XYZ((item.Value[i].Location as LocationPoint).Point.X, (item.Value[i].Location as LocationPoint).Point.Y - offset, (item.Value[i].Location as LocationPoint).Point.Z);

                                    ElementTransformUtils.MoveElement(doc, item.Value[i].Id, a - (item.Value[i].Location as LocationPoint).Point);
                                }
                            }
                        }
                        else
                        {
                            var vector = (item.Value[1].Location as LocationPoint).Point - (item.Value[0].Location as LocationPoint).Point;
                            if (vector.Y > 0)
                            {
                                if (isDown)
                                {
                                    XYZ a = new XYZ((item.Value[i].Location as LocationPoint).Point.X, (item.Value[i].Location as LocationPoint).Point.Y + offset, (item.Value[i].Location as LocationPoint).Point.Z);

                                    ElementTransformUtils.MoveElement(doc, item.Value[i].Id, a - (item.Value[i].Location as LocationPoint).Point);
                                }
                                else
                                {
                                    XYZ a = new XYZ((item.Value[i].Location as LocationPoint).Point.X, (item.Value[i].Location as LocationPoint).Point.Y - offset, (item.Value[i].Location as LocationPoint).Point.Z);

                                    ElementTransformUtils.MoveElement(doc, item.Value[i].Id, a - (item.Value[i].Location as LocationPoint).Point);
                                }
                            }
                            else
                            {
                                if (isDown)
                                {
                                    XYZ a = new XYZ((item.Value[i].Location as LocationPoint).Point.X, (item.Value[i].Location as LocationPoint).Point.Y - offset, (item.Value[i].Location as LocationPoint).Point.Z);

                                    ElementTransformUtils.MoveElement(doc, item.Value[i].Id, a - (item.Value[i].Location as LocationPoint).Point);
                                }
                                else
                                {
                                    XYZ a = new XYZ((item.Value[i].Location as LocationPoint).Point.X, (item.Value[i].Location as LocationPoint).Point.Y + offset, (item.Value[i].Location as LocationPoint).Point.Z);

                                    ElementTransformUtils.MoveElement(doc, item.Value[i].Id, a - (item.Value[i].Location as LocationPoint).Point);
                                }
                            }
                        }
                    }
                }
                else
                {
                    List<Connector> lstCon = GetListConnector(item.Value.FirstOrDefault().MEPModel.ConnectorManager.Connectors);
                    if (lstCon?.Count > 0)
                    {
                        var vector = lstCon[1].Origin - lstCon[0].Origin;
                        if (vector.Y > 0)
                        {
                            XYZ a = new XYZ((item.Value.FirstOrDefault().Location as LocationPoint).Point.X, (item.Value.FirstOrDefault().Location as LocationPoint).Point.Y + offset, (item.Value.FirstOrDefault().Location as LocationPoint).Point.Z);

                            ElementTransformUtils.MoveElement(doc, item.Value.FirstOrDefault().Id, a - (item.Value.FirstOrDefault().Location as LocationPoint).Point);
                        }
                        else
                        {
                            XYZ a = new XYZ((item.Value.FirstOrDefault().Location as LocationPoint).Point.X, (item.Value.FirstOrDefault().Location as LocationPoint).Point.Y - offset, (item.Value.FirstOrDefault().Location as LocationPoint).Point.Z);

                            ElementTransformUtils.MoveElement(doc, item.Value.FirstOrDefault().Id, a - (item.Value.FirstOrDefault().Location as LocationPoint).Point);
                        }
                    }
                }
            }
        }

        public static List<Connector> GetListConnector(ConnectorSet connectors)
        {
            List<Connector> connects = new List<Connector>();
            foreach (Connector c in connectors)
            {
                connects.Add(c);
            }
            return connects;
        }

        public static Dictionary<int, List<FamilyInstance>> GetDicElbow(List<FamilyInstance> lstElbow)
        {
            Dictionary<int, List<FamilyInstance>> dic = new Dictionary<int, List<FamilyInstance>>();
            int j = 1;
            for (int i = 0; i < lstElbow.Count; i++)
            {
                List<FamilyInstance> lst = new List<FamilyInstance>();
                lst.Add(lstElbow[i]);

                if (i != lstElbow.Count - 1)
                    lst.Add(lstElbow[i + 1]);

                dic.Add(j, lst);

                j++;

                if (lstElbow.Count != i + 1)
                    i = i + 1;
            }
            return dic;
        }

        /// <summary>
        /// set builtin parameter value
        /// </summary>
        public static bool SetBuiltinParameterValue(Element elem, BuiltInParameter paramId, object value)
        {
            Parameter param = elem.get_Parameter(paramId);
            return SetParameterValue(param, value);
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
    }
}