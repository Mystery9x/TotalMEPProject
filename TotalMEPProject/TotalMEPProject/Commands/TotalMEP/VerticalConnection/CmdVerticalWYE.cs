using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TotalMEPProject.SelectionFilters;
using TotalMEPProject.UI.TotalMEPUI;
using TotalMEPProject.Ultis;

namespace TotalMEPProject.Commands.TotalMEP
{
    [Transaction(TransactionMode.Manual)]
    public class CmdVerticalWYE : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Global.UIApp = commandData.Application;
            Global.RVTApp = commandData.Application.Application;
            Global.UIDoc = commandData.Application.ActiveUIDocument;
            Global.AppCreation = commandData.Application.Application.Create;

            bool isContinue = true;
            while (isContinue)
            {
                using (Transaction trans = new Transaction(Global.UIDoc.Document))
                {
                    trans.Start("VerticalConnection");
                    VerticalConnectionForm verticalConnectionForm = new VerticalConnectionForm();
                    if (verticalConnectionForm.DialogResult != System.Windows.Forms.DialogResult.OK)
                        break;

                    Pipe mainPipe = MEPUtilscs.PickPipe(Global.UIDoc, new PipeSelectionFilter(), "Please pick main pipe");
                    if (mainPipe == null)
                        break;

                    if (verticalConnectionForm.IsType1)
                    {
                    }
                    else
                    {
                    }
                }
            }
            return Result.Succeeded;
        }

        ///// <summary>
        ///// Kết nối ống theo TH1 (C1-TH1)
        ///// </summary>
        ///// <param name="doc"></param>
        ///// <param name="pipe1"></param>
        ///// <param name="pipe2"></param>
        ///// <param name="orgPoint"></param>
        ///// <param name="selectedHubmode"></param>
        //public static void CreateVerticalConnectionTH1(Document doc, ref Pipe pipe1, Pipe pipe2, XYZ orgPoint, ref string errMessage)
        //{
        //    int selectedHubmode = 1;
        //    try
        //    {
        //        Pipe pipeSplit = null;
        //        if (pipe1 != null && pipe2 != null && pipe1.Location is LocationCurve && pipe2.Location is LocationCurve)
        //        {
        //            bool isSucces = Common.DetermindPipeVertcalAndHorizontal(pipe1, pipe2, out MEPCurve pipeVertical, out MEPCurve pipeHorizontal);

        //            if (isSucces)
        //            {
        //                FamilySymbol typeWye = Common.GetSymbolSeted(doc, pipeHorizontal, RoutingPreferenceRuleGroupType.Junctions);

        //                FamilySymbol typeElbow = Common.GetSymbolSeted(doc, pipeVertical, RoutingPreferenceRuleGroupType.Elbows);
        //                if (typeWye == null || typeElbow == null)
        //                {
        //                    IO.ShowWarning(Define.ERR_NO_SET_ELBOW_FOR_PIPETYPE, "Waring");
        //                    return;
        //                }

        //                if (!Common.GetInforConnector(doc, typeWye, out int IdConSt, out int IdConEnd, out int IdConTee))
        //                {
        //                    return;
        //                }

        //                Common.MovePipeToCenter3(doc, pipeVertical, pipeHorizontal);

        //                XYZ locationWye = Common.GetPointIntersecNotInXYPlane(pipeVertical, pipeHorizontal);
        //                if (locationWye == null)
        //                {
        //                    return;
        //                }

        //                Connector conHor1 = ConnectorUtils.GetConnectorNearest(locationWye, pipeHorizontal, out Connector conHor2);

        //                FamilyInstance fittingWye = doc.Create.NewFamilyInstance(locationWye, typeWye, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

        //                Connector conSt = null;

        //                Connector conEnd = null;

        //                Connector conNhanhWye = null;

        //                if (fittingWye != null)
        //                {
        //                    conSt = fittingWye.MEPModel?.ConnectorManager.Lookup(IdConSt);

        //                    conEnd = fittingWye.MEPModel?.ConnectorManager.Lookup(IdConEnd);

        //                    conNhanhWye = fittingWye.MEPModel?.ConnectorManager.Lookup(IdConTee);

        //                    conSt.Radius = pipeVertical.Diameter / 2;
        //                    conEnd.Radius = pipeVertical.Diameter / 2;
        //                    conNhanhWye.Radius = pipeHorizontal.Diameter / 2;
        //                }

        //                doc.Regenerate();

        //                if (fittingWye == null)
        //                {
        //                    return;
        //                }

        //                Common.GetInformationConectorWye(fittingWye, null, out conSt, out conEnd, out conNhanhWye);

        //                if (conSt == null || conEnd == null || conNhanhWye == null)
        //                {
        //                    return;
        //                }

        //                Line axisRotate = ((LocationCurve)pipeVertical.Location).Curve as Line;

        //                Common.RotateLine(doc, fittingWye, axisRotate);

        //                doc.Regenerate();

        //                Line axisY = Line.CreateUnbound(locationWye, XYZ.BasisY);
        //                double angleFlip = Common.GetAngleFlipFitting2(doc, fittingWye, true);

        //                fittingWye.Location.Rotate(axisY, -angleFlip);

        //                double angleFlip2 = Common.GetAngleFlipFitting(doc, fittingWye, true);

        //                ParameterUtils.SetValueParameterByName(fittingWye, "Angle", angleFlip2);

        //                doc.Regenerate();

        //                double angle = Common.GetAngleBetweenTwoPipe(locationWye, conHor2.Origin, fittingWye);

        //                Line axisZ = Line.CreateUnbound(locationWye, XYZ.BasisZ);

        //                fittingWye.Location.Rotate(axisZ, -angle);

        //                doc.Regenerate();
        //                conNhanhWye = fittingWye.MEPModel.ConnectorManager.Lookup(IdConTee);

        //                XYZ vectorCheo = (conNhanhWye.Origin - ((LocationPoint)fittingWye.Location).Point).Normalize();

        //                double l1 = 0;
        //                double l2 = 0;
        //                Parameter paraL2 = fittingWye.LookupParameter("L2");
        //                if (paraL2 != null && paraL2.StorageType == StorageType.Double)
        //                    l2 = paraL2.AsDouble();

        //                conNhanhWye = fittingWye.MEPModel.ConnectorManager.Lookup(IdConTee);

        //                XYZ endPoint = Common.GetPointOnVector(conNhanhWye.Origin, vectorCheo, 500 / 304.8);
        //                Pipe pipeConnection = Pipe.Create(doc, pipeHorizontal.GetTypeId(), pipeHorizontal.ReferenceLevel.Id, conNhanhWye, endPoint);

        //                ParameterUtils.SetValueParameterByBuiltIn(pipeConnection, BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM, pipeHorizontal.MEPSystem.GetTypeId());

        //                XYZ tranlation = endPoint - conHor1.Origin;

        //                ICollection<ElementId> elementIds = ElementTransformUtils.CopyElement(doc, pipeHorizontal.Id, tranlation);

        //                Pipe pipeHorizontalCopy = doc.GetElement(elementIds.FirstOrDefault()) as Pipe;

        //                ConnectorUtils.GetConnectorClosedTo(pipeConnection.ConnectorManager, pipeHorizontalCopy.ConnectorManager, out Connector con1, out Connector con2);

        //                FamilyInstance elbow = null;
        //                if (con1 != null && con2 != null)
        //                    elbow = doc.Create.NewElbowFitting(con1, con2);

        //                doc.Delete(new List<ElementId> { pipeHorizontalCopy.Id });
        //                doc.Regenerate();
        //                ConnectorUtils.GetConnectorClosedTo(pipeConnection.ConnectorManager, elbow.MEPModel.ConnectorManager, out con1, out con2);

        //                Parameter paraL1 = elbow.LookupParameter("L1");
        //                if (paraL1 != null && paraL1.StorageType == StorageType.Double)
        //                    l1 = paraL1.AsDouble();

        //                double lenghtPipeConnection = l1 + l2;

        //                if (Common.IsEqual(lenghtPipeConnection, 0))
        //                    lenghtPipeConnection = 5;

        //                endPoint = Common.GetPointOnVector(conNhanhWye.Origin, vectorCheo, lenghtPipeConnection);

        //                Common.ResetLocation(pipeConnection, conNhanhWye.Origin, endPoint);
        //                doc.Regenerate();
        //                XYZ tranlation2 = con1.Origin - con2.Origin;

        //                ElementTransformUtils.MoveElement(doc, elbow.Id, tranlation2);
        //                doc.Regenerate();
        //                Connector conElbow = ConnectorUtils.GetConnectorNotConnnected(elbow.MEPModel.ConnectorManager);

        //                conHor1 = ConnectorUtils.GetConnectorNearest(conElbow.Origin, pipeHorizontal, out conHor2);

        //                XYZ vectorMoveZ1 = (conHor1.Origin.Z > conHor2.Origin.Z || Common.IsEqual(conHor1.Origin.Z, conHor2.Origin.Z)) ? XYZ.BasisZ : XYZ.BasisZ.Negate();

        //                XYZ newPoint2 = new XYZ(conElbow.Origin.X, conElbow.Origin.Y, conHor2.Origin.Z);
        //                double distance = newPoint2.DistanceTo(conHor2.Origin);
        //                double slope = Math.Round((double)ParameterUtils.GetValueParameterByBuilt(pipeHorizontal, BuiltInParameter.RBS_PIPE_SLOPE), 7);
        //                newPoint2 = Common.GetPointOnVector(newPoint2, vectorMoveZ1, slope * distance);

        //                //Ngắt kết nối với fitting
        //                Connector conStPipeHor = pipeHorizontal.ConnectorManager.Lookup(0);
        //                Connector conEndPipeHor = pipeHorizontal.ConnectorManager.Lookup(1);

        //                FamilyInstance elbowFittingConnectedHor = null;
        //                if (conStPipeHor != null && conEndPipeHor != null)
        //                {
        //                    if (conStPipeHor.IsConnected || conEndPipeHor.IsConnected)
        //                    {
        //                        Connector connectorIsConnecting = (conStPipeHor.IsConnected) ? conStPipeHor : conEndPipeHor;

        //                        elbowFittingConnectedHor = Common.GetFittingConnected(connectorIsConnecting);
        //                        if (elbowFittingConnectedHor != null)
        //                        {
        //                            ConnectorUtils.GetConnectorClosedTo(pipeHorizontal.ConnectorManager, elbowFittingConnectedHor.MEPModel.ConnectorManager, out Connector conPipeHor1, out Connector conFitHor);
        //                            if (conPipeHor1 != null && conFitHor != null && conPipeHor1.IsConnectedTo(conFitHor))
        //                                conPipeHor1.DisconnectFrom(conFitHor);
        //                        }
        //                    }
        //                }
        //                doc.Regenerate();
        //                Common.ResetLocation(pipeHorizontal, newPoint2, conHor2.Origin);

        //                // Join lại fitting
        //                if (elbowFittingConnectedHor != null)
        //                {
        //                    ConnectorUtils.GetConnectorClosedTo(pipeHorizontal.ConnectorManager, elbowFittingConnectedHor.MEPModel.ConnectorManager, out Connector conPipeHor1, out Connector conFitHor);
        //                    if (conPipeHor1 != null && conFitHor != null && !conPipeHor1.IsConnected && !conFitHor.IsConnected)
        //                    {
        //                        elbowFittingConnectedHor.Location.Move(conPipeHor1.Origin - conFitHor.Origin);

        //                        conPipeHor1.ConnectTo(conFitHor);
        //                    }
        //                }

        //                doc.Regenerate();

        //                XYZ tranlation3 = newPoint2 - conElbow.Origin;

        //                ElementTransformUtils.MoveElements(doc, new List<ElementId> { elbow.Id, pipeConnection.Id, fittingWye.Id }, tranlation3);

        //                ConnectorUtils.GetConnectorClosedTo(pipeHorizontal.ConnectorManager, elbow.MEPModel.ConnectorManager, out con1, out con2);
        //                if (con1 != null && con2 != null)
        //                    con1.ConnectTo(con2);

        //                //// Cắt ống đứng thành các ống phù hợp
        //                if (!Common.SpilitPipeVertical(doc, pipeVertical, fittingWye, ref pipeSplit))
        //                {
        //                    return;
        //                }

        //                doc.Regenerate();
        //                pipe1 = (pipeSplit == null) ? pipeVertical : Common.GetNextPipe(pipeVertical, pipeSplit, orgPoint);
        //            }
        //        }
        //    }
        //    catch (Exception)
        //    {
        //        errMessage = Define.ERR_ANGLE_ELBOW;
        //        return;
        //    }
        //}
    }
}