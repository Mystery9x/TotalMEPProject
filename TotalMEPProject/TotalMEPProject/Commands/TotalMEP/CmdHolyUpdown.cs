﻿using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml;
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

        public static double eleOld = 0;

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
                    if (!App.isApply)
                    {
                        string saveValue = string.Empty;
                        StorageUtility.SetValue(Global.UIDoc.Document.ProjectInformation, StorageUtility.m_MEP_HoLyUpDown_Guild, StorageUtility.m_MEP_HoLyUpDown, typeof(string), saveValue);
                        App.isApply = true;
                    }
                    else
                    {
                        object valueEntity = StorageUtility.GetValue(Global.UIDoc.Document.ProjectInformation, schema, StorageUtility.m_MEP_HoLyUpDown, typeof(string));

                        if (valueEntity != null)
                        {
                            var str = valueEntity as string;

                            ReadData(str);
                        }
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
                double offsetApply = UnitUtils.ConvertToInternalUnits(Math.Abs(App.m_HolyUpDownForm.UpStepValue), DisplayUnitType.DUT_MILLIMETERS);// Math.Abs(App.m_HolyUpDownForm.Distance) * Common.mmToFT;
                if (App.m_HolyUpDownForm.Distance < 0)
                {
                    offsetMain *= -1;
                }

                if (App.m_HolyUpDownForm.UpStepValue < 0)
                {
                    offsetApply *= -1;
                }

                bool isDeleteWarning = false;
                Transaction tran = new Transaction(Global.UIDoc.Document, "Holy");

                tran.Start();

                FailureHandlingOptions fhOpts = tran.GetFailureHandlingOptions();

                Dictionary<MEPCurve, List<FamilyInstance>> data = new Dictionary<MEPCurve, List<FamilyInstance>>();
                Dictionary<MEPCurve, List<FamilyInstance>> data_tap = new Dictionary<MEPCurve, List<FamilyInstance>>();
                Dictionary<ElementId, double> dicParaElevationBegin = new Dictionary<ElementId, double>();
                foreach (MEPCurve mepCurve in meps)
                {
                    var eleOld = GetBuiltInParameterValue(mepCurve, BuiltInParameter.RBS_OFFSET_PARAM)
                                    != null ? (double)GetBuiltInParameterValue(mepCurve, BuiltInParameter.RBS_OFFSET_PARAM) : double.MinValue;

                    dicParaElevationBegin.Add(mepCurve.Id, eleOld);
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
                            if (!App.m_HolyUpDownForm.NotApply)
                            {
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
                        double offsetApp = offsetApply;
                        if (dataMep == null)
                        {
                            dataMep = new Data(mepCurve);
                            var curve = (mepCurve.Location as LocationCurve).Curve;
                            dataMep._Start = curve.GetEndPoint(0);
                            dataMep._End = curve.GetEndPoint(1);
                            dataMep._OldOffset = 0;
                            dataMep._OldOffsetApply = offsetApp;
                            _Datas.Add(dataMep);

                            data_exist.Add(mepCurve, false);
                        }
                        else
                        {
                            data_exist.Add(mepCurve, true);
                        }
                        if (dataMep._Start != null && dataMep._End != null)
                        {
                            offset += dataMep._OldOffset;
                            dataMep._OldOffset = offset;

                            dataMep._OldOffsetApply = offsetApp;
                            XYZ p0New = new XYZ();
                            XYZ p1New = new XYZ();
                            if (!App.m_HolyUpDownForm.NotApply)
                            {
                                if ((mepCurve is CableTray || mepCurve is Conduit) && App.m_HolyUpDownForm.Elbow90)
                                {
                                    double offset1 = dicParaElevationBegin[mepCurve.Id] + offsetMain;
                                    SetBuiltinParameterValue(mepCurve, BuiltInParameter.RBS_OFFSET_PARAM, offset1);

                                    Global.UIDoc.Document.Regenerate();

                                    p0New = (mepCurve.Location as LocationCurve).Curve.GetEndPoint(0);
                                    p1New = (mepCurve.Location as LocationCurve).Curve.GetEndPoint(1);
                                }
                                else
                                {
                                    p0New = OffsetZ(dataMep._Start, offset);
                                    p1New = OffsetZ(dataMep._End, offset);

                                    (mepCurve.Location as LocationCurve).Curve = Line.CreateBound(p0New, p1New);
                                }

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
                    }

                    if (App.m_HolyUpDownForm.NotApply)
                    {
                        isDeleteWarning = true;
                        var lstSel = GetSelectedMEP();

                        List<Autodesk.Revit.DB.Mechanical.Duct> allDuctInModel = new List<Autodesk.Revit.DB.Mechanical.Duct>();
                        List<FamilyInstance> allDuctFittingInModel = new List<FamilyInstance>();
                        List<Pipe> allPipeInModel = new List<Pipe>();
                        List<FamilyInstance> allPipeFittingInModel = new List<FamilyInstance>();

                        if (lstSel.Any(item => item is Autodesk.Revit.DB.Mechanical.Duct))
                        {
                            allDuctInModel = new FilteredElementCollector(Global.UIDoc.Document).OfCategory(BuiltInCategory.OST_DuctCurves)
                                                                                        .OfClass(typeof(Autodesk.Revit.DB.Mechanical.Duct))
                                                                                        .Cast<Autodesk.Revit.DB.Mechanical.Duct>()
                                                                                        .ToList();

                            allDuctFittingInModel = new FilteredElementCollector(Global.UIDoc.Document).OfCategory(BuiltInCategory.OST_DuctFitting)
                                                                                                           .OfClass(typeof(FamilyInstance))
                                                                                                           .Cast<FamilyInstance>()
                                                                                                           .ToList();
                        }

                        if (lstSel.Any(item => item is Pipe))
                        {
                            allPipeInModel = new FilteredElementCollector(Global.UIDoc.Document).OfCategory(BuiltInCategory.OST_PipeCurves)
                                                                                    .OfClass(typeof(Autodesk.Revit.DB.Plumbing.Pipe))
                                                                                    .Cast<Autodesk.Revit.DB.Plumbing.Pipe>()
                                                                                    .ToList();

                            allPipeFittingInModel = new FilteredElementCollector(Global.UIDoc.Document).OfCategory(BuiltInCategory.OST_PipeFitting)
                                                                                                           .OfClass(typeof(FamilyInstance))
                                                                                                           .Cast<FamilyInstance>()
                                                                                                       .ToList();
                        }

                        foreach (MEPCurve mepCurve in lstSel)
                        {
                            MEPCurve processMepCurve = mepCurve;

                            if (mepCurve is Autodesk.Revit.DB.Mechanical.Duct duct)
                            {
                                SourceDuctData ductData = new SourceDuctData(mepCurve, allDuctInModel, allDuctFittingInModel);
                                processMepCurve = ductData.ProcessDuct != null ? ductData.ProcessDuct : mepCurve;
                            }

                            if (mepCurve is Autodesk.Revit.DB.Plumbing.Pipe pipe)
                            {
                                SourcePipeData pipeData = new SourcePipeData(mepCurve, allPipeInModel, allPipeFittingInModel);
                                processMepCurve = pipeData.ProcessPipe != null ? pipeData.ProcessPipe : mepCurve;
                            }

                            double dOldOffsetParamVal = GetBuiltInParameterValue(processMepCurve, BuiltInParameter.RBS_OFFSET_PARAM)
                                                != null ? (double)GetBuiltInParameterValue(processMepCurve, BuiltInParameter.RBS_OFFSET_PARAM) : double.MinValue;

                            if (dOldOffsetParamVal == double.MinValue)
                                continue;

                            if (dicParaElevationBegin.ContainsKey(mepCurve.Id))
                            {
                                var dataMep = _Datas.Find(item => item._MEPCurve.Id == mepCurve.Id);
                                if (dataMep == null)
                                    continue;
                                if (mepCurve is Autodesk.Revit.DB.Mechanical.Duct)
                                {
                                    double offset = dOldOffsetParamVal + dataMep._OldOffsetApply;
                                    SetBuiltinParameterValue(processMepCurve, BuiltInParameter.RBS_OFFSET_PARAM, offset);
                                }
                                else
                                {
                                    double offset = dicParaElevationBegin[mepCurve.Id] + dataMep._OldOffsetApply;
                                    SetBuiltinParameterValue(mepCurve, BuiltInParameter.RBS_OFFSET_PARAM, offset);
                                }
                            }
                        }

                        App.isApply = false;
                    }
                    else
                    {
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
                                if (mepCurve is CableTray || mepCurve is Conduit)
                                    dataMep.isCapleTray = true;
                                else
                                {
                                    Global.UIDoc.Document.Delete(dataMep._Vertical01.Id);
                                    dataMep._Vertical01 = null;
                                }
                            }

                            if (dataMep._Vertical02 != null && dataMep._Vertical02.IsValidObject == true)
                            {
                                if (mepCurve is CableTray || mepCurve is Conduit)
                                    dataMep.isCapleTray = true;
                                else
                                {
                                    Global.UIDoc.Document.Delete(dataMep._Vertical02.Id);
                                    dataMep._Vertical02 = null;
                                }
                            }

                            if (dataMep._Elbow01 != null && dataMep._Elbow01.IsValidObject == true)
                            {
                                if (mepCurve is CableTray || mepCurve is Conduit)
                                    dataMep.isCapleTray = true;
                                else
                                {
                                    Global.UIDoc.Document.Delete(dataMep._Elbow01.Id);
                                    dataMep._Elbow01 = null;
                                }
                            }

                            if (dataMep._Elbow02 != null && dataMep._Elbow02.IsValidObject == true)
                            {
                                if (mepCurve is CableTray || mepCurve is Conduit)
                                    dataMep.isCapleTray = true;
                                else
                                {
                                    Global.UIDoc.Document.Delete(dataMep._Elbow02.Id);
                                    dataMep._Elbow02 = null;
                                }
                            }

                            if (dataMep._Elbow03 != null && dataMep._Elbow03.IsValidObject == true)
                            {
                                if (mepCurve is CableTray || mepCurve is Conduit)
                                    dataMep.isCapleTray = true;
                                else
                                {
                                    Global.UIDoc.Document.Delete(dataMep._Elbow03.Id);
                                    dataMep._Elbow03 = null;
                                }
                            }

                            if (dataMep._Elbow04 != null && dataMep._Elbow04.IsValidObject == true)
                            {
                                if (mepCurve is CableTray || mepCurve is Conduit)
                                    dataMep.isCapleTray = true;
                                else
                                {
                                    Global.UIDoc.Document.Delete(dataMep._Elbow04.Id);
                                    dataMep._Elbow04 = null;
                                }
                            }
                        }

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
                                    if ((mepCurve is CableTray || mepCurve is Conduit) && dataMep.isCapleTray)
                                    {
                                        double offset = dicParaElevationBegin[mepCurve.Id] + offsetMain;
                                        SetBuiltinParameterValue(mepCurve, BuiltInParameter.RBS_OFFSET_PARAM, offset);
                                    }
                                    else
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
                                }
                                else
                                {
                                    if (mepCurve is CableTray && dataMep.isCapleTray)
                                    {
                                        double offset = dicParaElevationBegin[mepCurve.Id] + offsetMain;
                                        SetBuiltinParameterValue(mepCurve, BuiltInParameter.RBS_OFFSET_PARAM, offset);
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
                }

                var values = WriteXML();

                bool result = StorageUtility.AddEntity(Global.UIDoc.Document.ProjectInformation, StorageUtility.m_MEP_HoLyUpDown_Guild, StorageUtility.m_MEP_HoLyUpDown, values);

                GetInfoWarning supWarning = new GetInfoWarning(isDeleteWarning);
                fhOpts.SetFailuresPreprocessor(supWarning);
                tran.SetFailureHandlingOptions(fhOpts);

                tran.Commit(fhOpts);

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
                    if (element != null && element is FamilyInstance && element.MEPModel != null)
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

                        attribute = node.Attributes["OldOffsetApply"];
                        if (attribute != null && attribute.Value != null)
                        {
                            dataMep._OldOffsetApply = ToDouble(attribute.Value);
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

                attribute = xmlDoc.CreateAttribute("OldOffsetApply");
                if (data._OldOffsetApply != double.NaN)
                    attribute.Value = data._OldOffsetApply.ToString();
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

                List<Autodesk.Revit.DB.Mechanical.Duct> allDuctInModel = new List<Autodesk.Revit.DB.Mechanical.Duct>();
                List<FamilyInstance> allDuctFittingInModel = new List<FamilyInstance>();
                List<Pipe> allPipeInModel = new List<Pipe>();
                List<FamilyInstance> allPipeFittingInModel = new List<FamilyInstance>();

                if (mepCurveSelects.Any(item => item is Autodesk.Revit.DB.Mechanical.Duct))
                {
                    allDuctInModel = new FilteredElementCollector(Global.UIDoc.Document).OfCategory(BuiltInCategory.OST_DuctCurves)
                                                                                .OfClass(typeof(Autodesk.Revit.DB.Mechanical.Duct))
                                                                                .Cast<Autodesk.Revit.DB.Mechanical.Duct>()
                                                                                .ToList();

                    allDuctFittingInModel = new FilteredElementCollector(Global.UIDoc.Document).OfCategory(BuiltInCategory.OST_DuctFitting)
                                                                                                   .OfClass(typeof(FamilyInstance))
                                                                                                   .Cast<FamilyInstance>()
                                                                                                   .ToList();
                }

                if (mepCurveSelects.Any(item => item is Pipe))
                {
                    allPipeInModel = new FilteredElementCollector(Global.UIDoc.Document).OfCategory(BuiltInCategory.OST_PipeCurves)
                                                                            .OfClass(typeof(Autodesk.Revit.DB.Plumbing.Pipe))
                                                                            .Cast<Autodesk.Revit.DB.Plumbing.Pipe>()
                                                                            .ToList();

                    allPipeFittingInModel = new FilteredElementCollector(Global.UIDoc.Document).OfCategory(BuiltInCategory.OST_PipeFitting)
                                                                                                   .OfClass(typeof(FamilyInstance))
                                                                                                   .Cast<FamilyInstance>()
                                                                                               .ToList();
                }

                double dStepValue = UnitUtils.ConvertToInternalUnits(Math.Abs(App.m_HolyUpDownForm.UpStepValue), DisplayUnitType.DUT_MILLIMETERS);

                foreach (MEPCurve mepCurve in mepCurveSelects)
                {
                    MEPCurve processMepCurve = mepCurve;

                    if (mepCurve is Autodesk.Revit.DB.Mechanical.Duct duct)
                    {
                        List<ElementId> allElementIds = allDuctFittingInModel.Select(item => item.Id).ToList();

                        // Create a Outline, uses a minimum and maximum XYZ point to initialize the outline.
                        Outline myOutLn = CreateOutLineFromBoundingBox(duct);
                        if (myOutLn == null || myOutLn.IsEmpty)
                            continue;

                        // Create a BoundingBoxIntersects filter with this Outline
                        BoundingBoxIntersectsFilter filter = new BoundingBoxIntersectsFilter(myOutLn);

                        FilteredElementCollector collector = new FilteredElementCollector(Global.UIDoc.Document, allElementIds);

                        List<FamilyInstance> allFittingConnected = collector.WherePasses(filter).Cast<FamilyInstance>().ToList();

                        SourceDuctData ductData = new SourceDuctData(mepCurve, allDuctInModel, allFittingConnected);
                        processMepCurve = ductData.ProcessDuct != null ? ductData.ProcessDuct : mepCurve;
                    }

                    if (mepCurve is Autodesk.Revit.DB.Plumbing.Pipe pipe)
                    {
                        List<ElementId> allElementIds = allPipeFittingInModel.Select(item => item.Id).ToList();

                        // Create a Outline, uses a minimum and maximum XYZ point to initialize the outline.
                        Outline myOutLn = CreateOutLineFromBoundingBox(pipe);
                        if (myOutLn == null || myOutLn.IsEmpty)
                            continue;

                        // Create a BoundingBoxIntersects filter with this Outline
                        BoundingBoxIntersectsFilter filter = new BoundingBoxIntersectsFilter(myOutLn);

                        FilteredElementCollector collector = new FilteredElementCollector(Global.UIDoc.Document, allElementIds);

                        List<FamilyInstance> allFittingConnected = collector.WherePasses(filter).Cast<FamilyInstance>().ToList();

                        SourcePipeData pipeData = new SourcePipeData(mepCurve, allPipeInModel, allFittingConnected);
                        processMepCurve = pipeData.ProcessPipe != null ? pipeData.ProcessPipe : mepCurve;
                    }

                    double dOldOffsetParamVal = GetBuiltInParameterValue(processMepCurve, BuiltInParameter.RBS_OFFSET_PARAM)
                                                != null ? (double)GetBuiltInParameterValue(processMepCurve, BuiltInParameter.RBS_OFFSET_PARAM) : double.MinValue;

                    if (dOldOffsetParamVal == double.MinValue)
                        continue;
                    using (Transaction reTrans = new Transaction(Global.UIDoc.Document, "HOLYUPDOWN_UPSTEP"))
                    {
                        if (reTrans.Start() == TransactionStatus.Started)
                        {
                            FailureHandlingOptions fhOpts = reTrans.GetFailureHandlingOptions();
                            try
                            {
                                double dNewOffsetParamVal = downStep == false ? dOldOffsetParamVal + dStepValue : dOldOffsetParamVal - dStepValue;
                                SetBuiltinParameterValue(processMepCurve, BuiltInParameter.RBS_OFFSET_PARAM, dNewOffsetParamVal);

                                GetInfoWarning supWarning = new GetInfoWarning(true);
                                fhOpts.SetFailuresPreprocessor(supWarning);
                                reTrans.SetFailureHandlingOptions(fhOpts);
                                reTrans.Commit(fhOpts);
                            }
                            catch (Exception)
                            {
                                reTrans.RollBack();
                            }
                        }
                    }

                    App.isApply = false;
                }
            }
            catch (Exception)
            {
            }
            return Result.Cancelled;
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
            XYZ vectorMinMax = (max - min).Normalize();
            retVal = new Outline(min + vectorMinMax.Negate() * 0.01, max + vectorMinMax * 0.01);
            return retVal;
        }

        public static MEPCurve GetMepSame(List<FamilyInstance> lstElbow)
        {
            MEPCurve retVal = null;
            Dictionary<FamilyInstance, List<ElementId>> dic = new Dictionary<FamilyInstance, List<ElementId>>();
            foreach (var elbow in lstElbow)
            {
                Connector c = null;

                var loc = elbow.Location as LocationPoint;
                if (loc == null)
                    continue;

                List<ElementId> pipesId = new List<ElementId>();
                foreach (Connector con in elbow.MEPModel.ConnectorManager.Connectors)
                {
                    var pipeCon = new FilteredElementCollector(Global.UIDoc.Document).OfClass(typeof(MEPCurve)).Cast<MEPCurve>().Where(x => x.Location as LocationCurve != null
                                                                                                                                && (Common.IsEqual((x.Location as LocationCurve).Curve.GetEndPoint(0), con.Origin)
                                                                                                                                || Common.IsEqual((x.Location as LocationCurve).Curve.GetEndPoint(1), con.Origin)))
                                                                                                       .FirstOrDefault();
                    if (pipeCon != null)
                        pipesId.Add(pipeCon.Id);
                }

                dic.Add(elbow, pipesId);
            }

            Dictionary<ElementId, List<FamilyInstance>> dicSamePipe = new Dictionary<ElementId, List<FamilyInstance>>();
            List<FamilyInstance> lstElbowNoSame = new List<FamilyInstance>();

            foreach (var item in dic)
            {
                foreach (var item1 in dic)
                {
                    if (item.Key.Id.IntegerValue == item1.Key.Id.IntegerValue)
                        continue;

                    var samePipeId = item.Value.Intersect(item1.Value).FirstOrDefault();

                    if (samePipeId != null)
                    {
                        List<FamilyInstance> lstElbowSame = new List<FamilyInstance>();
                        retVal = Global.UIDoc.Document.GetElement(samePipeId) as MEPCurve;
                        return retVal;
                    }
                }
            }

            return retVal;
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

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        public static Result Run_UpdownElbowControl1(bool downElbowControl = false)
        {
            try
            {
                using (Transaction reTrans = new Transaction(Global.UIDoc.Document, "HOLY_UPDOWN_ELBOW_CONTROL"))
                {
                    reTrans.Start();

                    try
                    {
                        if (App.m_HolyUpDownForm == null)
                            return Result.Cancelled;

                        var selElbows = GetSelectedElbow();
                        if (selElbows == null || selElbows.Count == 0)
                            return Result.Cancelled;

                        List<MEPCurveData> filterMEPCurve = new FilteredElementCollector(Global.UIDoc.Document).OfClass(typeof(MEPCurve))
                                                                                    .Cast<MEPCurve>()
                                                                                    .Select(item => new MEPCurveData(item))
                                                                                    .ToList();

                        List<ElbowControlData> elbowControlDatas = new List<ElbowControlData>();
                        foreach (FamilyInstance elbow in selElbows)
                        {
                            elbowControlDatas.Add(new ElbowControlData(elbow, filterMEPCurve));
                        }

                        List<ElbowControlData> validElbow = elbowControlDatas.Where(item => item.MEPCurveConnects.Count > 0).ToList();

                        var filterElbows = FilterElbow(validElbow).ToList();

                        double offset = App.m_HolyUpDownForm.UpElbowStepValue / 304.8;

                        foreach (List<ElbowControlData> elbows in filterElbows)
                        {
                            List<ElbowControlData> processElbow = elbows.Distinct(new DistinctItemComparer()).ToList();

                            if (processElbow.Count > 1)
                            {
                                var twoElbow = GetTwoPointFurthest(processElbow);
                                ElbowControlData elbow_1 = twoElbow.Item1;
                                ElbowControlData elbow_2 = twoElbow.Item2;

                                List<FamilyInstance> lstElbow = new List<FamilyInstance>();
                                lstElbow.Add(elbow_1.ElbowMain);
                                lstElbow.Add(elbow_2.ElbowMain);
                                double slope = 0;
                                var mepSame = GetMepSame(lstElbow);
                                if (mepSame != null)
                                {
                                    if (mepSame.get_Parameter(BuiltInParameter.RBS_PIPE_SLOPE) != null)
                                        slope = GetBuiltInParameterValue(mepSame, BuiltInParameter.RBS_PIPE_SLOPE);
                                    else if (mepSame.get_Parameter(BuiltInParameter.RBS_DUCT_SLOPE) != null)
                                        slope = GetBuiltInParameterValue(mepSame, BuiltInParameter.RBS_DUCT_SLOPE);
                                }

                                XYZ p1_1 = elbow_1.LocationPoint + elbow_1.MEPCurveDataMain.Direction.Normalize() * offset;
                                XYZ p1_2 = elbow_1.LocationPoint + elbow_1.MEPCurveDataMain.Direction.Normalize().Negate() * offset;

                                XYZ p2_1 = elbow_2.LocationPoint + elbow_2.MEPCurveDataMain.Direction.Normalize() * offset;
                                XYZ p2_2 = elbow_2.LocationPoint + elbow_2.MEPCurveDataMain.Direction.Normalize().Negate() * offset;

                                double distance_1 = p1_1.DistanceTo(p2_1);
                                double distance_2 = p1_1.DistanceTo(p2_2);
                                double distance_3 = p1_2.DistanceTo(p2_1);
                                double distance_4 = p1_2.DistanceTo(p2_2);

                                List<double> distances = new List<double>() { distance_1, distance_2, distance_3, distance_4 };
                                double minVal = distances.Min();
                                double maxVal = distances.Max();

                                if (slope < 0.03)
                                {
                                    if (downElbowControl)
                                    {
                                        if (minVal == distance_1)
                                        {
                                            ElementTransformUtils.MoveElement(Global.UIDoc.Document, elbow_1.ElbowMain.Id, elbow_1.MEPCurveDataMain.Direction.Normalize() * offset);
                                            ElementTransformUtils.MoveElement(Global.UIDoc.Document, elbow_2.ElbowMain.Id, elbow_2.MEPCurveDataMain.Direction.Normalize() * offset);
                                        }
                                        else if (minVal == distance_2)
                                        {
                                            ElementTransformUtils.MoveElement(Global.UIDoc.Document, elbow_1.ElbowMain.Id, elbow_1.MEPCurveDataMain.Direction.Normalize() * offset);
                                            ElementTransformUtils.MoveElement(Global.UIDoc.Document, elbow_2.ElbowMain.Id, elbow_2.MEPCurveDataMain.Direction.Normalize().Negate() * offset);
                                        }
                                        else if (minVal == distance_3)
                                        {
                                            ElementTransformUtils.MoveElement(Global.UIDoc.Document, elbow_1.ElbowMain.Id, elbow_1.MEPCurveDataMain.Direction.Normalize().Negate() * offset);
                                            ElementTransformUtils.MoveElement(Global.UIDoc.Document, elbow_2.ElbowMain.Id, elbow_2.MEPCurveDataMain.Direction.Normalize() * offset);
                                        }
                                        else
                                        {
                                            ElementTransformUtils.MoveElement(Global.UIDoc.Document, elbow_1.ElbowMain.Id, elbow_1.MEPCurveDataMain.Direction.Normalize().Negate() * offset);
                                            ElementTransformUtils.MoveElement(Global.UIDoc.Document, elbow_2.ElbowMain.Id, elbow_2.MEPCurveDataMain.Direction.Normalize().Negate() * offset);
                                        }
                                    }
                                    else
                                    {
                                        if (maxVal == distance_1)
                                        {
                                            ElementTransformUtils.MoveElement(Global.UIDoc.Document, elbow_1.ElbowMain.Id, elbow_1.MEPCurveDataMain.Direction.Normalize() * offset);
                                            ElementTransformUtils.MoveElement(Global.UIDoc.Document, elbow_2.ElbowMain.Id, elbow_2.MEPCurveDataMain.Direction.Normalize() * offset);
                                        }
                                        else if (maxVal == distance_2)
                                        {
                                            ElementTransformUtils.MoveElement(Global.UIDoc.Document, elbow_1.ElbowMain.Id, elbow_1.MEPCurveDataMain.Direction.Normalize() * offset);
                                            ElementTransformUtils.MoveElement(Global.UIDoc.Document, elbow_2.ElbowMain.Id, elbow_2.MEPCurveDataMain.Direction.Normalize().Negate() * offset);
                                        }
                                        else if (maxVal == distance_3)
                                        {
                                            ElementTransformUtils.MoveElement(Global.UIDoc.Document, elbow_1.ElbowMain.Id, elbow_1.MEPCurveDataMain.Direction.Normalize().Negate() * offset);
                                            ElementTransformUtils.MoveElement(Global.UIDoc.Document, elbow_2.ElbowMain.Id, elbow_2.MEPCurveDataMain.Direction.Normalize() * offset);
                                        }
                                        else
                                        {
                                            ElementTransformUtils.MoveElement(Global.UIDoc.Document, elbow_1.ElbowMain.Id, elbow_1.MEPCurveDataMain.Direction.Normalize().Negate() * offset);
                                            ElementTransformUtils.MoveElement(Global.UIDoc.Document, elbow_2.ElbowMain.Id, elbow_2.MEPCurveDataMain.Direction.Normalize().Negate() * offset);
                                        }
                                    }
                                }
                                else
                                {
                                    if (downElbowControl)
                                    {
                                        if (minVal == distance_1)
                                            ElementTransformUtils.MoveElement(Global.UIDoc.Document, elbow_1.ElbowMain.Id, elbow_1.MEPCurveDataMain.Direction.Normalize() * offset);
                                        else if (minVal == distance_2)
                                            ElementTransformUtils.MoveElement(Global.UIDoc.Document, elbow_1.ElbowMain.Id, elbow_1.MEPCurveDataMain.Direction.Normalize() * offset);
                                        else if (minVal == distance_3)
                                            ElementTransformUtils.MoveElement(Global.UIDoc.Document, elbow_1.ElbowMain.Id, elbow_1.MEPCurveDataMain.Direction.Normalize().Negate() * offset);
                                        else
                                            ElementTransformUtils.MoveElement(Global.UIDoc.Document, elbow_1.ElbowMain.Id, elbow_1.MEPCurveDataMain.Direction.Normalize().Negate() * offset);
                                    }
                                    else
                                    {
                                        if (maxVal == distance_1)
                                            ElementTransformUtils.MoveElement(Global.UIDoc.Document, elbow_1.ElbowMain.Id, elbow_1.MEPCurveDataMain.Direction.Normalize() * offset);
                                        else if (maxVal == distance_2)
                                            ElementTransformUtils.MoveElement(Global.UIDoc.Document, elbow_1.ElbowMain.Id, elbow_1.MEPCurveDataMain.Direction.Normalize() * offset);
                                        else if (maxVal == distance_3)
                                            ElementTransformUtils.MoveElement(Global.UIDoc.Document, elbow_1.ElbowMain.Id, elbow_1.MEPCurveDataMain.Direction.Normalize().Negate() * offset);
                                        else
                                            ElementTransformUtils.MoveElement(Global.UIDoc.Document, elbow_1.ElbowMain.Id, elbow_1.MEPCurveDataMain.Direction.Normalize().Negate() * offset);
                                    }
                                }
                            }
                            else if (processElbow.Count > 0)
                            {
                                ElbowControlData elbow = processElbow[0];

                                if (downElbowControl)
                                {
                                    ElementTransformUtils.MoveElement(Global.UIDoc.Document, elbow.ElbowMain.Id, -elbow.MEPCurveDataMain.Direction.Normalize() * offset);
                                }
                                else
                                {
                                    ElementTransformUtils.MoveElement(Global.UIDoc.Document, elbow.ElbowMain.Id, elbow.MEPCurveDataMain.Direction.Normalize() * offset);
                                }
                            }
                        }

                        reTrans.Commit();
                    }
                    catch (Exception)
                    {
                        reTrans.RollBack();
                    }
                }
            }
            catch (Exception)
            { }
            return Result.Cancelled;
        }

        private static IEnumerable<List<ElbowControlData>> FilterElbow(List<ElbowControlData> elbowControlDatas)
        {
            List<ElbowControlData> processElbows = new List<ElbowControlData>(elbowControlDatas);
            List<ElbowControlData> processElbows_1 = new List<ElbowControlData>(elbowControlDatas);

            for (int i = 0; i < processElbows.Count; i++)
            {
                ElbowControlData elbowMain = processElbows[i];
                List<ElbowControlData> retVal = new List<ElbowControlData>();

                processElbows_1.Remove(elbowMain);
                var elbowDirectionDuplicate = processElbows_1.Where(item => PointOnLineSegment(elbowMain.MEPCurveDataMain.StartPointXtend,
                                                                                               elbowMain.MEPCurveDataMain.EndPointXtend,
                                                                                               item.MEPCurveDataMain.StartPoint) && PointOnLineSegment(elbowMain.MEPCurveDataMain.StartPointXtend,
                                                                                                                                                       elbowMain.MEPCurveDataMain.EndPointXtend,
                                                                                                                                                       item.MEPCurveDataMain.EndPoint))
                                                             .ToList();

                if (elbowDirectionDuplicate.Count >= 0)
                {
                    processElbows.Remove(elbowMain);
                    i--;

                    retVal.Add(elbowMain);
                    retVal.AddRange(elbowDirectionDuplicate);
                    elbowDirectionDuplicate.ForEach(item =>
                    {
                        processElbows_1.Remove(item);
                        processElbows.Remove(item);
                    });

                    yield return retVal;
                }
            }
        }

        public static bool PointOnLineSegment(XYZ pt1, XYZ pt2, XYZ pt, double epsilon = 1e-6)
        {
            if (pt.X - Math.Max(pt1.X, pt2.X) > epsilon ||
                Math.Min(pt1.X, pt2.X) - pt.X > epsilon ||
                pt.Y - Math.Max(pt1.Y, pt2.Y) > epsilon ||
                Math.Min(pt1.Y, pt2.Y) - pt.Y > epsilon)
                return false;

            if (Math.Abs(pt2.X - pt1.X) < epsilon)
                return Math.Abs(pt1.X - pt.X) < epsilon || Math.Abs(pt2.X - pt.X) < epsilon;
            if (Math.Abs(pt2.Y - pt1.Y) < epsilon)
                return Math.Abs(pt1.Y - pt.Y) < epsilon || Math.Abs(pt2.Y - pt.Y) < epsilon;

            var x = pt1.X + (pt.Y - pt1.Y) * (pt2.X - pt1.X) / (pt2.Y - pt1.Y);
            var y = pt1.Y + (pt.X - pt1.X) * (pt2.Y - pt1.Y) / (pt2.X - pt1.X);

            return Math.Abs(pt.X - x) < epsilon || Math.Abs(pt.Y - y) < epsilon;
        }

        // Check two point
        public static Tuple<ElbowControlData, ElbowControlData> GetTwoPointFurthest(List<ElbowControlData> elbowControlDatas)
        {
            try
            {
                double max = -1;

                List<Tuple<double, Tuple<ElbowControlData, ElbowControlData>>> temp = new List<Tuple<double, Tuple<ElbowControlData, ElbowControlData>>>();
                for (int i = 0; i < elbowControlDatas.Count - 1; i++)
                {
                    for (int j = i + 1; j < elbowControlDatas.Count; j++)
                    {
                        XYZ p1 = new XYZ(elbowControlDatas[i].LocationPoint.X, elbowControlDatas[i].LocationPoint.Y, 0);
                        XYZ p2 = new XYZ(elbowControlDatas[j].LocationPoint.X, elbowControlDatas[j].LocationPoint.Y, 0);
                        double distance = p1.DistanceTo(p2);
                        Tuple<double, Tuple<ElbowControlData, ElbowControlData>> tuple = new Tuple<double, Tuple<ElbowControlData, ElbowControlData>>(distance, new Tuple<ElbowControlData, ElbowControlData>(elbowControlDatas[i], elbowControlDatas[j]));
                        temp.Add(tuple);
                    }
                }

                return temp.OrderByDescending(item => item.Item1).FirstOrDefault().Item2;
            }
            catch (Exception)
            { }

            return null;
        }

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    }

    public class GetInfoWarning : IFailuresPreprocessor
    {
        public static bool _isApply;

        public GetInfoWarning(bool isApply)
        {
            _isApply = isApply;
        }

        public FailureProcessingResult PreprocessFailures(FailuresAccessor failuresAccessor)
        {
            if (_isApply)
            {
                var messages = failuresAccessor.GetFailureMessages();
                if (messages.Count() > 0)
                {
                    foreach (FailureMessageAccessor message in messages)
                    {
                        var lstId = message.GetFailingElementIds();
                        failuresAccessor.DeleteWarning(message);
                    }
                }
            }

            return FailureProcessingResult.Continue;
        }
    }

    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    public class ElbowControlData
    {
        private FamilyInstance m_elbow = null;
        private List<MEPCurve> m_MEPCurveConnects = new List<MEPCurve>();
        private Connector m_connector1 = null;
        private Connector m_connector2 = null;
        private MEPCurveData m_MEPCurveDataMain = null;
        private MEPCurve m_MEPCurveDataOutOfThread = null;
        private Tuple<double, double> m_linearEquations = new Tuple<double, double>(double.MinValue, double.MinValue);
        private XYZ m_locationPoint = null;

        public FamilyInstance ElbowMain { get => m_elbow; set => m_elbow = value; }
        public Connector Connector1 { get => m_connector1; set => m_connector1 = value; }
        public Connector Connector2 { get => m_connector2; set => m_connector2 = value; }
        public Tuple<double, double> LinearEquations { get => m_linearEquations; set => m_linearEquations = value; }

        public List<MEPCurveData> AllMEPCurveDataInModels = new List<MEPCurveData>();

        public List<MEPCurveData> MEPCurveConnects = new List<MEPCurveData>();

        public List<MEPCurve> MEPCurveConnectsOutOfThread = new List<MEPCurve>();
        public MEPCurveData MEPCurveDataMain { get => m_MEPCurveDataMain; set => m_MEPCurveDataMain = value; }

        public MEPCurve MEPCurveDataOutOfThread { get => m_MEPCurveDataOutOfThread; set => m_MEPCurveDataOutOfThread = value; }
        public XYZ LocationPoint { get => m_locationPoint; set => m_locationPoint = value; }

        public ElbowControlData(FamilyInstance elbow, List<MEPCurveData> allMEPCurveModels)
        {
            ElbowMain = elbow;
            AllMEPCurveDataInModels = allMEPCurveModels;
            GetConnector();
            GetLocationPoint();
            GetLinearEquations();
            GetMEPCurveConnects();
        }

        public void GetConnector()
        {
            if (m_elbow != null)
            {
                if (m_elbow.MEPModel != null)
                {
                    List<Connector> cntorOfElbow = GetConnectorFromElbow(m_elbow.MEPModel.ConnectorManager.Connectors);
                    if (cntorOfElbow != null && cntorOfElbow.Count == 2)
                    {
                        Connector1 = cntorOfElbow[0];
                        Connector2 = cntorOfElbow[1];
                    }
                }
            }
        }

        public void GetLinearEquations()
        {
            if (Connector1 != null && Connector2 != null)
            {
                XYZ vector = (Connector1.Origin - Connector2.Origin).Normalize();
                double u = vector.X;
                double v = vector.Y;

                if (u < 0)
                {
                    u *= -1;
                    v *= -1;
                }
                LinearEquations = new Tuple<double, double>(u, v);
            }
        }

        public void GetMEPCurveConnects()
        {
            var filterMEPCurve = AllMEPCurveDataInModels.Where(item => item.Connector1 != null && item.Connector2 != null)
                                                    .Where(item => item.Connector1.IsConnectedTo(Connector1)
                                                                   || item.Connector1.IsConnectedTo(Connector2)
                                                                   || item.Connector2.IsConnectedTo(Connector1)
                                                                   || item.Connector2.IsConnectedTo(Connector2))
                                                    .ToList();

            MEPCurveConnects = filterMEPCurve.Where(item => item.Slope <= 0.00021).ToList();
            MEPCurveConnectsOutOfThread = filterMEPCurve.Where(item => item.Slope > 0.00021).Select(item => item.MEPCurveMain).ToList();
            if (MEPCurveConnects.Count > 0)
            {
                MEPCurveDataMain = MEPCurveConnects[0];
            }
            if (MEPCurveConnectsOutOfThread.Count > 0)
            {
                MEPCurveDataOutOfThread = MEPCurveConnectsOutOfThread[0];
            }
        }

        public void GetLocationPoint()
        {
            if (m_elbow != null)
            {
                LocationPoint = (m_elbow.Location as LocationPoint).Point;
            }
        }

        public List<Connector> GetConnectorFromElbow(ConnectorSet connectorSet)
        {
            try
            {
                List<Connector> retVal = new List<Connector>();
                foreach (Connector connector in connectorSet)
                {
                    retVal.Add(connector);
                }
                return retVal;
            }
            catch (Exception)
            { }
            return new List<Connector>();
        }
    }

    public class MEPCurveData
    {
        private MEPCurve m_MEPCurve = null;
        private double m_slope = double.MaxValue;
        private XYZ m_startPoint = null;
        private XYZ m_endPoint = null;

        private XYZ m_direction = null;
        private XYZ m_directionFlatten = null;
        private Connector m_connector1 = null;
        private Connector m_connector2 = null;

        private List<Connector> m_connectors = new List<Connector>();

        public MEPCurve MEPCurveMain { get => m_MEPCurve; set => m_MEPCurve = value; }
        public double Slope { get => m_slope; set => m_slope = value; }
        public XYZ StartPoint { get => m_startPoint; set => m_startPoint = value; }
        public XYZ EndPoint { get => m_endPoint; set => m_endPoint = value; }
        public XYZ Direction { get => m_direction; set => m_direction = value; }
        public XYZ DirectionFlatten { get => m_directionFlatten; set => m_directionFlatten = value; }

        public XYZ StartPointXtend
        {
            get
            {
                if (StartPoint != null)
                {
                    return StartPoint - Direction.Normalize() * 100;
                }
                return StartPoint;
            }
            private set
            {
                StartPoint = value;
            }
        }

        public XYZ EndPointXtend
        {
            get
            {
                if (EndPoint != null)
                {
                    return EndPoint + Direction.Normalize() * 100;
                }
                return EndPoint;
            }
            private set
            {
                EndPoint = value;
            }
        }

        public Connector Connector1 { get => m_connector1; set => m_connector1 = value; }
        public Connector Connector2 { get => m_connector2; set => m_connector2 = value; }

        public List<Connector> Connectors { get => m_connectors; set => m_connectors = value; }

        public MEPCurveData(MEPCurve mEPCurve)
        {
            MEPCurveMain = mEPCurve;
            Initialize();
        }

        private void Initialize()
        {
            if (MEPCurveMain != null && MEPCurveMain.Location as LocationCurve != null)
            {
                StartPoint = (MEPCurveMain.Location as LocationCurve).Curve.GetEndPoint(0);
                EndPoint = (MEPCurveMain.Location as LocationCurve).Curve.GetEndPoint(1);

                GetConnector();

                try
                {
                    XYZ directionTemp = ((MEPCurveMain.Location as LocationCurve).Curve as Line).Direction;

                    Slope = 1 - Math.Sin(directionTemp.AngleTo(XYZ.BasisZ));
                    Direction = EndPoint - StartPoint;
                    DirectionFlatten = new XYZ(Direction.X, Direction.Y, 0);
                }
                catch (Exception)
                { }
            }
        }

        public void GetConnector()
        {
            if (m_MEPCurve != null)
            {
                if (m_MEPCurve.ConnectorManager != null)
                {
                    List<Connector> cntorOfElbow = GetConnectorFromPipe(m_MEPCurve.ConnectorManager.Connectors);
                    Connectors = GetAllConnectorFromMEPCurve(m_MEPCurve.ConnectorManager.Connectors);
                    if (cntorOfElbow != null && cntorOfElbow.Count == 2)
                    {
                        Connector1 = cntorOfElbow[0];
                        Connector2 = cntorOfElbow[1];
                    }
                }
            }
        }

        public List<Connector> GetConnectorFromPipe(ConnectorSet connectorSet)
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

        public List<Connector> GetAllConnectorFromMEPCurve(ConnectorSet connectorSet)
        {
            try
            {
                List<Connector> retVal = new List<Connector>();
                foreach (Connector connector in connectorSet)
                {
                    retVal.Add(connector);
                }
                return retVal;
            }
            catch (Exception)
            { }
            return new List<Connector>();
        }

        public dynamic GetBuiltInParameterValue(Element elem, BuiltInParameter paramId)
        {
            if (elem != null)
            {
                Parameter parameter = elem.get_Parameter(paramId);
                return GetParameterValue(parameter);
            }
            return null;
        }

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
    }

    public class SourceDuctData
    {
        private Autodesk.Revit.DB.Mechanical.Duct m_mainDuct = null;
        public Autodesk.Revit.DB.Mechanical.Duct MainDuct { get => m_mainDuct; set => m_mainDuct = value; }

        private List<Autodesk.Revit.DB.Mechanical.Duct> m_allDucts = new List<Autodesk.Revit.DB.Mechanical.Duct>();
        public List<Autodesk.Revit.DB.Mechanical.Duct> AllDucts { get => m_allDucts; set => m_allDucts = value; }

        private List<FamilyInstance> m_allDuctFittings = new List<FamilyInstance>();
        public List<FamilyInstance> AllDuctFittings { get => m_allDuctFittings; set => m_allDuctFittings = value; }

        private Autodesk.Revit.DB.Mechanical.Duct m_processDuct = null;
        public Autodesk.Revit.DB.Mechanical.Duct ProcessDuct { get => m_processDuct; set => m_processDuct = value; }

        private List<Connector> m_connectorsOfDuct = new List<Connector>();
        public List<Connector> ConnectorOfDutct { get => m_connectorsOfDuct; set => m_connectorsOfDuct = value; }

        private SourceDuctType m_ductType = SourceDuctType.Nothing;
        public SourceDuctType DuctType { get => m_ductType; set => m_ductType = value; }

        private List<SourceTagData> m_tags = new List<SourceTagData>();
        public List<SourceTagData> Tags { get => m_tags; set => m_tags = value; }
        public Connector FirstConnector { get; set; }
        public Connector SecondConnector { get; set; }

        public SourceDuctData(MEPCurve mepCurve,
                              List<Autodesk.Revit.DB.Mechanical.Duct> allDuctInModel,
                              List<FamilyInstance> allDuctFittingInModel)
        {
            if (mepCurve is Autodesk.Revit.DB.Mechanical.Duct duct
                && allDuctInModel != null
                && allDuctInModel.Count > 0
                && allDuctFittingInModel != null
                && allDuctFittingInModel.Count > 0)
            {
                MainDuct = duct;
                AllDucts = allDuctInModel;
                AllDuctFittings = allDuctFittingInModel;
                Tags = allDuctFittingInModel.Select(item => new SourceTagData(item, AllDucts)).ToList();
                Initialize();
            }
        }

        private void Initialize()
        {
            GetAllConnectorOfDuct();
            List<SourceTagData> tagsConnect = new List<SourceTagData>();
            if (ConnectorOfDutct.Count > 2)
                DuctType = SourceDuctType.MainDuct;
            else if (ConnectorOfDutct.Count == 2)
            {
                tagsConnect = GetTagsConnect();
                if (tagsConnect.Count <= 0)
                    DuctType = SourceDuctType.MainDuct;
                else
                {
                    SourceTagData firstTagConnect = tagsConnect[0];
                    if (FirstConnector.IsConnectedTo(firstTagConnect.SecondConnector) || SecondConnector.IsConnectedTo(firstTagConnect.SecondConnector))
                    {
                        DuctType = SourceDuctType.BranchDuct;
                    }
                    else
                        DuctType = SourceDuctType.MainDuct;
                }
            }
            if (DuctType == SourceDuctType.BranchDuct)
            {
                SourceTagData firstTagConnect = tagsConnect[0];
                var ductConnectTag = firstTagConnect.DuctsConnect.Where(item => item.Id != MainDuct.Id).FirstOrDefault();

                if (ductConnectTag != null) { ProcessDuct = ductConnectTag; }
            }
            else
                ProcessDuct = MainDuct;
        }

        public void GetAllConnectorOfDuct()
        {
            if (MainDuct != null)
            {
                if (MainDuct.ConnectorManager != null)
                {
                    List<Connector> cntorOfDuct = GetConnectors(MainDuct.ConnectorManager.Connectors);
                    ConnectorOfDutct = cntorOfDuct;

                    FirstConnector = cntorOfDuct.Where(item => item.ConnectorType == ConnectorType.End).FirstOrDefault();
                    SecondConnector = cntorOfDuct.Where(item => item.ConnectorType == ConnectorType.End).LastOrDefault();
                }
            }
        }

        public List<Connector> GetConnectors(ConnectorSet connectorSet, bool filter = false)
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

        private List<SourceTagData> GetTagsConnect()
        {
            try
            {
                List<SourceTagData> filterTags = Tags.Where(item => item.FirstConnector != null && item.SecondConnector != null)
                                                        .Where(item => item.FirstConnector.IsConnectedTo(FirstConnector)
                                                                       || item.FirstConnector.IsConnectedTo(SecondConnector)
                                                                       || item.SecondConnector.IsConnectedTo(FirstConnector)
                                                                       || item.SecondConnector.IsConnectedTo(SecondConnector))
                                                        .ToList();
                return filterTags;
            }
            catch (Exception)
            { }
            return new List<SourceTagData>();
        }
    }

    public class SourceTagData
    {
        public FamilyInstance TagData { get; set; }
        public Connector FirstConnector { get; set; }
        public Connector SecondConnector { get; set; }
        public List<Autodesk.Revit.DB.Mechanical.Duct> AllDuct { get; set; }

        public List<Autodesk.Revit.DB.Mechanical.Duct> DuctsConnect { get; set; }

        public SourceTagData(FamilyInstance tagData, List<Autodesk.Revit.DB.Mechanical.Duct> allDuctInModel)
        {
            TagData = tagData;
            AllDuct = allDuctInModel;
            if (TagData != null && TagData.MEPModel != null && TagData.MEPModel.ConnectorManager != null)
            {
                List<Connector> cntOfTags = GetConnectors(TagData.MEPModel.ConnectorManager.Connectors, true);
                if (cntOfTags.Count >= 2)
                {
                    FirstConnector = cntOfTags.Where(item => item.Id == 1).FirstOrDefault();
                    SecondConnector = cntOfTags.Where(item => item.Id == 2).FirstOrDefault();
                    GetDuctsConnect();
                }
            }
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

        private void GetDuctsConnect()
        {
            DuctsConnect = new List<Autodesk.Revit.DB.Mechanical.Duct>();
            List<MEPCurveData> allDucts = AllDuct.Select(item => new MEPCurveData(item as MEPCurve)).ToList();

            List<MEPCurveData> filterMEPCurve = FirstConnector != null ? allDucts.Where(item => item.Connector1 != null && item.Connector2 != null).Where(item => item.Connectors.Any(item1 => item1.IsConnectedTo(FirstConnector) == true)).ToList() : new List<MEPCurveData>();

            if (filterMEPCurve.Count > 0)
                DuctsConnect = filterMEPCurve.Select(item => item.MEPCurveMain as Autodesk.Revit.DB.Mechanical.Duct).ToList();
        }
    }

    public enum SourceDuctType : int
    {
        Nothing = -1,
        MainDuct,
        BranchDuct
    }

    public class SourcePipeData
    {
        private Autodesk.Revit.DB.Plumbing.Pipe m_mainPipe = null;
        public Autodesk.Revit.DB.Plumbing.Pipe MainPipe { get => m_mainPipe; set => m_mainPipe = value; }

        private List<Autodesk.Revit.DB.Plumbing.Pipe> m_allPipes = new List<Autodesk.Revit.DB.Plumbing.Pipe>();
        public List<Autodesk.Revit.DB.Plumbing.Pipe> AllPipes { get => m_allPipes; set => m_allPipes = value; }

        private List<FamilyInstance> m_allPipeFittings = new List<FamilyInstance>();
        public List<FamilyInstance> AllPipeFittings { get => m_allPipeFittings; set => m_allPipeFittings = value; }

        private Autodesk.Revit.DB.Plumbing.Pipe m_processPipe = null;
        public Autodesk.Revit.DB.Plumbing.Pipe ProcessPipe { get => m_processPipe; set => m_processPipe = value; }

        private List<Connector> m_connectorsOfPipe = new List<Connector>();
        public List<Connector> ConnectorOfDutct { get => m_connectorsOfPipe; set => m_connectorsOfPipe = value; }

        private SourcePipeType m_PipeType = SourcePipeType.Nothing;
        public SourcePipeType PipeType { get => m_PipeType; set => m_PipeType = value; }

        private List<SourceTagPipeData> m_tags = new List<SourceTagPipeData>();
        public List<SourceTagPipeData> Tags { get => m_tags; set => m_tags = value; }
        public Connector FirstConnector { get; set; }
        public Connector SecondConnector { get; set; }

        public SourcePipeData(MEPCurve mepCurve,
                              List<Autodesk.Revit.DB.Plumbing.Pipe> allPipeInModel,
                              List<FamilyInstance> allPipeFittingInModel)
        {
            if (mepCurve is Autodesk.Revit.DB.Plumbing.Pipe Pipe
                && allPipeInModel != null
                && allPipeInModel.Count > 0
                && allPipeFittingInModel != null
                && allPipeFittingInModel.Count > 0)
            {
                MainPipe = Pipe;
                AllPipes = allPipeInModel;
                AllPipeFittings = allPipeFittingInModel;
                Tags = allPipeFittingInModel.Where(item => item.MEPModel != null && item.MEPModel is MechanicalFitting fitting && fitting.PartType != Autodesk.Revit.DB.PartType.Elbow).Select(item => new SourceTagPipeData(item, AllPipes)).ToList();
                Initialize();
            }
        }

        private void Initialize()
        {
            GetAllConnectorOfPipe();
            List<SourceTagPipeData> tagsConnect = new List<SourceTagPipeData>();
            if (ConnectorOfDutct.Count > 2)
                PipeType = SourcePipeType.MainPipe;
            else if (ConnectorOfDutct.Count == 2)
            {
                tagsConnect = GetTagsConnect();
                if (tagsConnect.Count <= 0)
                    PipeType = SourcePipeType.MainPipe;
                else
                {
                    SourceTagPipeData firstTagConnect = tagsConnect[0];
                    if (FirstConnector.IsConnectedTo(firstTagConnect.SecondConnector) || SecondConnector.IsConnectedTo(firstTagConnect.SecondConnector))
                    {
                        PipeType = SourcePipeType.BranchPipe;
                    }
                    else if (FirstConnector.IsConnectedTo(firstTagConnect.FirstConnector) || SecondConnector.IsConnectedTo(firstTagConnect.FirstConnector))
                    {
                        PipeType = SourcePipeType.BranchPipe;
                    }
                    else
                        PipeType = SourcePipeType.MainPipe;
                }
            }
            if (PipeType == SourcePipeType.BranchPipe)
            {
                SourceTagPipeData firstTagConnect = tagsConnect[0];
                var PipeConnectTag = firstTagConnect.PipesConnect.Where(item => item.Id != MainPipe.Id).FirstOrDefault();

                if (PipeConnectTag != null) { ProcessPipe = PipeConnectTag; }
            }
            else
                ProcessPipe = MainPipe;
        }

        public void GetAllConnectorOfPipe()
        {
            if (MainPipe != null)
            {
                if (MainPipe.ConnectorManager != null)
                {
                    List<Connector> cntorOfPipe = GetConnectors(MainPipe.ConnectorManager.Connectors);
                    ConnectorOfDutct = cntorOfPipe;

                    FirstConnector = cntorOfPipe.Where(item => item.ConnectorType == ConnectorType.End).FirstOrDefault();
                    SecondConnector = cntorOfPipe.Where(item => item.ConnectorType == ConnectorType.End).LastOrDefault();
                }
            }
        }

        public List<Connector> GetConnectors(ConnectorSet connectorSet, bool filter = false)
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

        private List<SourceTagPipeData> GetTagsConnect()
        {
            try
            {
                List<SourceTagPipeData> filterTags = Tags.Where(item => item.FirstConnector != null && item.SecondConnector != null)
                                                        .Where(item => item.FirstConnector.IsConnectedTo(FirstConnector)
                                                                       || item.FirstConnector.IsConnectedTo(SecondConnector)
                                                                       || item.SecondConnector.IsConnectedTo(FirstConnector)
                                                                       || item.SecondConnector.IsConnectedTo(SecondConnector))
                                                        .ToList();
                return filterTags;
            }
            catch (Exception)
            { }
            return new List<SourceTagPipeData>();
        }
    }

    public class SourceTagPipeData
    {
        public FamilyInstance TagData { get; set; }
        public Connector FirstConnector { get; set; }
        public Connector SecondConnector { get; set; }
        public List<Autodesk.Revit.DB.Plumbing.Pipe> AllPipe { get; set; }

        public List<Autodesk.Revit.DB.Plumbing.Pipe> PipesConnect { get; set; }

        public SourceTagPipeData(FamilyInstance tagData, List<Autodesk.Revit.DB.Plumbing.Pipe> allPipeInModel)
        {
            TagData = tagData;
            AllPipe = allPipeInModel;
            if (TagData != null && TagData.MEPModel != null && TagData.MEPModel.ConnectorManager != null)
            {
                List<Connector> cntOfTags = GetConnectors(TagData.MEPModel.ConnectorManager.Connectors, true);
                if (cntOfTags.Count >= 2)
                {
                    FirstConnector = cntOfTags.Where(item => item.Id == 1).FirstOrDefault();
                    SecondConnector = cntOfTags.Where(item => item.Id == 2).FirstOrDefault();
                    GetPipesConnect();
                }
            }
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

        private void GetPipesConnect()
        {
            PipesConnect = new List<Autodesk.Revit.DB.Plumbing.Pipe>();
            List<MEPCurveData> allPipes = AllPipe.Select(item => new MEPCurveData(item as MEPCurve)).ToList();

            List<MEPCurveData> filterMEPCurve = allPipes.Where(item => item.Connector1 != null && item.Connector2 != null).Where(item => item.Connectors.Any(item1 => item1.IsConnectedTo(FirstConnector) == true) || item.Connectors.Any(item1 => item1.IsConnectedTo(SecondConnector) == true)).ToList();

            if (filterMEPCurve.Count > 0)
                PipesConnect = filterMEPCurve.Select(item => item.MEPCurveMain as Autodesk.Revit.DB.Plumbing.Pipe).ToList();
        }
    }

    public enum SourcePipeType : int
    {
        Nothing = -1,
        MainPipe,
        BranchPipe
    }

    internal class DistinctItemComparer : IEqualityComparer<ElbowControlData>
    {
        public bool Equals(ElbowControlData x, ElbowControlData y)
        {
            return x.MEPCurveDataOutOfThread.Id == y.MEPCurveDataOutOfThread.Id;
        }

        public int GetHashCode(ElbowControlData obj)
        {
            return obj.MEPCurveDataOutOfThread.Id.GetHashCode();
        }
    }

    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
}