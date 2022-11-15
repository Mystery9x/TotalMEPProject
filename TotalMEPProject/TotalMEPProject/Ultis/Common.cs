using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Plumbing;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;

namespace TotalMEPProject.Ultis
{
    public static class Common
    {
        public static double mmToFT = 0.0032808399;
        private const double _eps = 1.0e-9;

        public static bool IsLessThan(double first, double second, double tolerance = 10e-5)
        {
            if (!IsEqual(first, second, tolerance))
                return first < second;

            return false;
        }

        public static bool IsLessOrEqualThan(double first, double second, double tolerance = 10e-5)
        {
            if (!IsEqual(first, second, tolerance))
                return first < second;

            return true;
        }

        public static bool IsGreaterOrEqualThan(double first, double second, double tolerance = 10e-5)
        {
            if (!IsEqual(first, second, tolerance))
                return first > second;

            return true;
        }

        public static bool IsGreaterThan(double first, double second, double tolerance = 10e-5)
        {
            if (!IsEqual(first, second, tolerance))
                return first > second;

            return false;
        }

        public static bool IsEqual(double first, double second, double tolerance = 10e-5)
        {
            double result = Math.Abs(first - second);
            return result < tolerance;
        }

        public static bool IsEqual(XYZ first, XYZ second)
        {
            return IsEqual(first.X, second.X)
                && IsEqual(first.Y, second.Y)
                && IsEqual(first.Z, second.Z);
        }

        /// <summary>
        /// determine if 2 vectors/ points are almost equal within a toelrance range
        /// </summary>
        /// <param name="first"></param>
        /// <param name="second"></param>
        /// <param name="tolerance"></param>
        /// <returns></returns>
        public static bool IsEqualPoint2D(XYZ first, XYZ second)
        {
            return IsEqual(first.X, second.X) &&
                   IsEqual(first.Y, second.Y);
        }

        public static bool IsParallel(XYZ p, XYZ q)
        {
            if (p.CrossProduct(q).IsZeroLength() == true)
                return true;

            var l = p.CrossProduct(q).GetLength();
            if (IsZero(l, 10e-5))
                return true;

            return false;
        }

        public static bool IsZero(double a, double tolerance)
        {
            return tolerance > Math.Abs(a);
        }

        public static bool IsEqual(double first, double second)
        {
            double result = Math.Abs(first - second);
            return result < 10e-5;
        }

        public static void NumberCheck(object sender, KeyPressEventArgs e, bool allowNegativeValue = false) // < 0
        {
            if (allowNegativeValue == false)
            {
                if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
                {
                    e.Handled = true;
                }

                // only allow one decimal point
                if ((e.KeyChar == '.') && ((sender as System.Windows.Forms.TextBox).Text.IndexOf('.') > -1))
                {
                    e.Handled = true;
                }
            }
            else
            {
                if (!char.IsControl(e.KeyChar) && (!char.IsDigit(e.KeyChar)) && (e.KeyChar != '.') && (e.KeyChar != '-'))
                    e.Handled = true;

                // only allow one decimal point
                if (e.KeyChar == '.' && (sender as System.Windows.Forms.TextBox).Text.IndexOf('.') > -1)
                    e.Handled = true;

                // only allow minus sign at the beginning
                if (e.KeyChar == '-' && (sender as System.Windows.Forms.TextBox).Text.IndexOf('-') > -1)
                    e.Handled = true;
            }
        }

        public static XYZ ToPoint2D(XYZ point3d, double z = 0)
        {
            return new XYZ(point3d.X, point3d.Y, z);
        }

        public static List<Connector> ToList(ConnectorSet connectors)
        {
            List<Connector> connects = new List<Connector>();
            foreach (Connector c in connectors)
            {
                connects.Add(c);
            }
            return connects;
        }

        /// <summary>
        /// Lấy điểm cách level 1 khoảng cách cho trước
        /// </summary>
        /// <param name="level"></param>
        /// <param name="point"></param>
        /// <param name="offsetFromLevel"></param>
        /// <returns></returns>
        public static XYZ GetPointOffsetFromLevel(Level level, XYZ point, double offsetFromLevel)
        {
            if (level == null || point == null)
                return null;

            offsetFromLevel = UnitUtils.ConvertToInternalUnits(offsetFromLevel, DisplayUnitType.DUT_MILLIMETERS);
            if (level == null || point == null)
                return point;

            return new XYZ(point.X, point.Y, level.Elevation + offsetFromLevel);
        }

        /// <summary>
        /// Di chuyển 2 ống về cùng tâm
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="pipe1"></param>
        /// <param name="pipe2"></param>
        /// <returns></returns>
        public static bool MovePipeToCenter(Document doc, MEPCurve pipe1, MEPCurve pipe2)
        {
            if (pipe1 != null && pipe2 != null && pipe1.Location is LocationCurve locationCurve1 && pipe2.Location is LocationCurve locationCurve2)
            {
                XYZ st1 = locationCurve1.Curve.GetEndPoint(0);
                st1 = new XYZ(st1.X, st1.Y, 0);

                XYZ end1 = locationCurve1.Curve.GetEndPoint(1);
                end1 = new XYZ(end1.X, end1.Y, 0);

                double lenght1 = st1.DistanceTo(end1);

                Line line1 = null;
                Line line2 = null;

                if (lenght1 > UnitUtils.ConvertToInternalUnits(1, DisplayUnitType.DUT_MILLIMETERS) && !st1.IsAlmostEqualTo(end1))
                    line1 = Line.CreateBound(st1, end1);

                XYZ st2 = locationCurve2.Curve.GetEndPoint(0);
                st2 = new XYZ(st2.X, st2.Y, 0);

                XYZ end2 = locationCurve2.Curve.GetEndPoint(1);
                end2 = new XYZ(end2.X, end2.Y, 0);

                double lenght2 = st2.DistanceTo(end2);

                if (lenght2 > UnitUtils.ConvertToInternalUnits(1, DisplayUnitType.DUT_MILLIMETERS) && !st2.IsAlmostEqualTo(end2))
                    line2 = Line.CreateBound(st2, end2);

                Line lineValid = (lenght1 > lenght2) ? line1 : line2;

                XYZ point = (lenght1 > lenght2) ? st2 : st1;

                if (lineValid != null && point != null)
                {
                    Line lineUnbound = Line.CreateUnbound(lineValid.Origin, lineValid.Direction);

                    IntersectionResult intersection = lineUnbound.Project(point);
                    if (intersection != null)
                    {
                        MEPCurve pipeMove = (lenght1 > lenght2) ? pipe1 : pipe2;

                        XYZ tranMove = point - intersection.XYZPoint;

                        //Move element
                        ElementTransformUtils.MoveElement(doc, pipeMove.Id, tranMove);
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// Xác định ống thẳng đứng và ống ngang
        /// </summary>
        /// <param name="pipe1"></param>
        /// <param name="pipe2"></param>
        /// <param name="pipeVertical"></param>
        /// <param name="pipeHorizontal"></param>
        /// <returns></returns>
        public static bool DetermindPipeVertcalAndHorizontal(MEPCurve pipe1, MEPCurve pipe2, out MEPCurve pipeVertical, out MEPCurve pipeHorizontal)
        {
            pipeVertical = pipe1;
            pipeHorizontal = pipe2;

            if (pipe1 != null && pipe2 != null && pipe1.Location is LocationCurve locationCurve1 && pipe2.Location is LocationCurve locationCurve2)
            {
                XYZ st1 = locationCurve1.Curve.GetEndPoint(0);
                st1 = new XYZ(st1.X, st1.Y, 0);

                XYZ end1 = locationCurve1.Curve.GetEndPoint(1);
                end1 = new XYZ(end1.X, end1.Y, 0);

                //Line line1 = null;
                //Line line2 = null;

                double lenght1 = st1.DistanceTo(end1);

                //if (!st1.IsAlmostEqualTo(end1))
                //    line1 = Line.CreateBound(st1, end1);

                XYZ st2 = locationCurve2.Curve.GetEndPoint(0);
                st2 = new XYZ(st2.X, st2.Y, 0);

                XYZ end2 = locationCurve2.Curve.GetEndPoint(1);
                end2 = new XYZ(end2.X, end2.Y, 0);

                double lenght2 = st2.DistanceTo(end2);

                //if (!st2.IsAlmostEqualTo(end2))
                //    line2 = Line.CreateBound(st2, end2);

                //double lenght1 = (line1 != null) ? line1.Length : 0;

                //double lenght2 = (line2 != null) ? line2.Length : 0;

                pipeVertical = (lenght1 < lenght2) ? pipe1 : pipe2;

                pipeHorizontal = (lenght1 > lenght2) ? pipe1 : pipe2;

                if (pipeHorizontal.Id.IntegerValue != pipeVertical.Id.IntegerValue)
                    return true;
            }

            return false;
        }

        /// <summary>
        /// Lấy 2 connector thuộc 2 ống khác nhau và gần nhau nhất
        /// </summary>
        /// <param name="connectorManager1"></param>
        /// <param name="connectorManager2"></param>
        /// <param name="con1"></param>
        /// <param name="con2"></param>
        public static void GetConnectorClosedTo(ConnectorManager connectorManager1, ConnectorManager connectorManager2, out Connector con1, out Connector con2)
        {
            con1 = null;
            con2 = null;

            if (connectorManager1 != null && connectorManager2 != null)

            {
                double distanceMin = double.MaxValue;

                foreach (Connector item1 in connectorManager1.Connectors)
                {
                    foreach (Connector item2 in connectorManager2.Connectors)
                    {
                        double distance = item1.Origin.DistanceTo(item2.Origin);
                        if (distance < distanceMin)
                        {
                            con1 = item1;
                            con2 = item2;
                            distanceMin = distance;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Tìm giao điểm của 2 pipe thẳng đứng và nằm ngang (không cùng trên 1 mặt phẳng)
        /// </summary>
        /// <param name="pipeVertical"></param>
        /// <param name="pipeHorizontal"></param>
        /// <returns></returns>
        public static XYZ GetPointIntersecNotInXYPlane(MEPCurve pipeVertical, MEPCurve pipeHorizontal)
        {
            if (pipeVertical != null && pipeHorizontal != null && pipeVertical.Location is LocationCurve && pipeHorizontal.Location is LocationCurve)
            {
                GetConnectorClosedTo(pipeVertical.ConnectorManager, pipeHorizontal.ConnectorManager, out Connector con1, out Connector con2);

                Line line = ((LocationCurve)pipeVertical.Location).Curve as Line;

                Line lineUnbound = Line.CreateUnbound(line.Origin, line.Direction);

                IntersectionResult intersection = lineUnbound.Project(con2.Origin);

                if (intersection != null)
                    return intersection.XYZPoint;
            }

            return null;
        }

        /// <summary>
        ///  sét giá trị parameter theo BuiltInParameter
        /// </summary>
        /// <param name="element"></param>
        /// <param name="builtIn"></param>
        /// <param name="valuePara"></param>
        /// <returns></returns>
        public static bool SetValueParameterByBuiltIn(Element element, BuiltInParameter builtIn, object valuePara)
        {
            if (element == null)
                return false;
            Parameter prm = element.get_Parameter(builtIn);
            if (prm != null && !prm.IsReadOnly)
            {
                if (prm.StorageType == StorageType.ElementId)
                {
                    prm.Set((ElementId)valuePara);

                    return true;
                }
                if (prm.StorageType == StorageType.Double)
                {
                    prm.Set((double)valuePara);

                    return true;
                }
                if (prm.StorageType == StorageType.Integer)
                {
                    prm.Set((int)valuePara);

                    return true;
                }
                if (prm.StorageType == StorageType.String)
                {
                    prm.Set((string)valuePara);

                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Lấy giá trị của paramter theo built in parameter
        /// </summary>
        /// <param name="element"></param>
        /// <param name="buintinparameter"></param>
        /// <returns></returns>
        public static object GetValueParameterByBuilt(Element element, BuiltInParameter buintinparameter)
        {
            if (element == null)
                return null;
            Parameter prm = element.get_Parameter(buintinparameter);
            if (prm != null)
            {
                if (prm.StorageType == StorageType.ElementId)
                {
                    return prm.AsElementId();
                }
                if (prm.StorageType == StorageType.Double)
                {
                    return prm.AsDouble();
                }
                if (prm.StorageType == StorageType.Integer)
                {
                    return prm.AsInteger();
                }
                if (prm.StorageType == StorageType.String)
                {
                    return prm.AsString();
                }
            }
            return null;
        }

        /// <summary>
        /// Lấy ra connector gần nhất và xa nhất với 1 điểm cho trước
        /// </summary>
        /// <param name="point"></param>
        /// <param name="pipe"></param>
        /// <param name="outFarest"></param>
        /// <returns></returns>
        public static Connector GetConnectorNearest(XYZ point, MEPCurve pipe, out Connector outFarest)
        {
            Connector retval = null;
            outFarest = null;

            if (point != null && pipe != null)
            {
                ConnectorManager connectorManager = pipe.ConnectorManager;

                double max = double.MaxValue;
                double min = double.MinValue;

                foreach (Connector item in connectorManager.Connectors)
                {
                    double distance = item.Origin.DistanceTo(point);

                    // lấy connector gần nhất
                    if (distance < max)
                    {
                        max = distance;
                        retval = item;
                    }
                    // lấy connector xa nhất
                    if (distance > min)
                    {
                        min = distance;
                        outFarest = item;
                    }
                }
            }

            return retval;
        }

        /// <summary>
        /// Lấy connector có cao độ cao hơn
        /// </summary>
        /// <param name="pipe"></param>
        /// <returns></returns>
        public static Connector GetConnectorValid(MEPCurve pipe, Level level)
        {
            if (pipe != null && pipe.Location is LocationCurve locationCurve)
            {
                double slope = 0;

                if (pipe is Pipe)
                    slope = Math.Round((double)GetValueParameterByBuilt(pipe, BuiltInParameter.RBS_PIPE_SLOPE), 5);
                else if (pipe is Autodesk.Revit.DB.Mechanical.Duct)
                    slope = Math.Round((double)GetValueParameterByBuilt(pipe, BuiltInParameter.RBS_DUCT_SLOPE), 5);

                XYZ pointSt = locationCurve.Curve.GetEndPoint(0);

                Connector conSt = GetConnectorNearest(pointSt, pipe, out Connector conEnd);

                if (Common.IsEqual(slope, 0))
                {
                    if (conEnd.IsConnected)
                        return conSt;

                    return conEnd;
                }
                else
                {
                    var levelPipe = pipe.ReferenceLevel;
                    if (levelPipe.Elevation < level.Elevation)
                    {
                        Connector retval = (conSt.Origin.Z > conEnd.Origin.Z) ? conSt : conEnd;
                        if (retval.IsConnected)
                            return (conSt.Origin.Z < conEnd.Origin.Z) ? conSt : conEnd;

                        return retval;
                    }
                    else
                    {
                        Connector retval = (conSt.Origin.Z > conEnd.Origin.Z) ? conEnd : conSt;
                        if (retval.IsConnected)
                            return (conSt.Origin.Z < conEnd.Origin.Z) ? conSt : conEnd;

                        return retval;
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// Check family siphon vertical or horizontal
        /// </summary>
        /// <param name="fittingSiphon"></param>
        /// <returns></returns>
        public static bool IsSiphonVertical(FamilyInstance fittingSiphon)
        {
            if (fittingSiphon != null && fittingSiphon.MEPModel != null)
            {
                Connector con1 = fittingSiphon.MEPModel.ConnectorManager.Lookup(1);

                Connector con2 = fittingSiphon.MEPModel.ConnectorManager.Lookup(2);

                if (con1 != null && con2 != null)
                {
                    if (IsEqual(con1.Origin.Z, con2.Origin.Z))
                        return false;
                }
            }

            return true;
        }

        /// <summary>
        ///  Tính góc giữa siphon và pipe nằm ngang
        /// </summary>
        /// <param name="pipe"></param>
        /// <param name="siphon"></param>
        /// <returns></returns>
        public static double GetAngleBetweenPipeAndSiphon(MEPCurve pipe, FamilyInstance siphon)
        {
            if (pipe != null && pipe.Location is LocationCurve location && siphon != null && siphon.MEPModel != null)
            {
                XYZ st1 = location.Curve.GetEndPoint(0);
                st1 = new XYZ(st1.X, st1.Y, 0);

                XYZ end1 = location.Curve.GetEndPoint(1);
                end1 = new XYZ(end1.X, end1.Y, 0);

                if (st1.IsAlmostEqualTo(end1))
                    return 0;

                ConnectorManager connectorManager = siphon.MEPModel.ConnectorManager;

                if (connectorManager == null)
                    return 0;

                Connector connectorS = null;
                Connector connectorE = null;

                if (pipe.Name.Equals(Define.PIPE_CAST_IRON))
                {
                    connectorS = connectorManager.Lookup(1);
                    connectorE = connectorManager.Lookup(2);
                }
                else
                {
                    connectorS = connectorManager.Lookup(2);
                    connectorE = connectorManager.Lookup(1);
                }

                if (connectorS != null && connectorE != null)

                {
                    XYZ conSt = new XYZ(connectorS.Origin.X, connectorS.Origin.Y, 0);

                    XYZ conEnd = new XYZ(connectorE.Origin.X, connectorE.Origin.Y, 0);

                    if (conSt.IsAlmostEqualTo(conEnd))
                        return 0;

                    XYZ poin1 = GetPointFarther(conSt, st1, end1, out XYZ poin2);

                    XYZ vec1 = (poin1 - poin2).Normalize();
                    XYZ vec2 = (conEnd - conSt).Normalize();

                    if (vec1 != null && vec2 != null)
                        return vec1.AngleOnPlaneTo(vec2, XYZ.BasisZ);
                }
            }
            return 0;
        }

        /// <summary>
        /// Lấy điểm xa hơn trong 2 điểm và trả ra điểm gần
        /// </summary>
        /// <param name="pointOrigin"></param>
        /// <param name="point1"></param>
        /// <param name="point2"></param>
        /// <returns></returns>
        private static XYZ GetPointFarther(XYZ pointOrigin, XYZ point1, XYZ point2, out XYZ pointNear)
        {
            XYZ retval = null;
            pointNear = null;
            if (pointOrigin != null && point1 != null && point2 != null)
            {
                double distance1 = pointOrigin.DistanceTo(point1);

                double distance2 = pointOrigin.DistanceTo(point2);

                if (distance1 > distance2)
                {
                    retval = point1;
                    pointNear = point2;
                }
                else
                {
                    retval = point2;
                    pointNear = point1;
                }
            }

            return retval;
        }

        /// <summary>
        /// Set giá trị cho paramter theo tên
        /// </summary>
        /// <param name="el"></param>
        /// <param name="parameterName"></param>
        /// <param name="valuePara"></param>
        /// <returns></returns>
        public static bool SetValueParameterByListName(Element el, object valuePara, params string[] parameterNames)
        {
            if (el == null || parameterNames.Length == 0)
                return false;
            foreach (string parameterName in parameterNames)
            {
                if (string.IsNullOrEmpty(parameterName))
                    continue;
                Parameter prm = el.LookupParameter(parameterName);
                if (prm != null && !prm.IsReadOnly)
                {
                    if (prm.StorageType == StorageType.ElementId)
                    {
                        prm.Set((ElementId)valuePara);

                        return true;
                    }
                    if (prm.StorageType == StorageType.Double)
                    {
                        prm.Set((double)valuePara);

                        return true;
                    }
                    if (prm.StorageType == StorageType.Integer)
                    {
                        prm.Set((int)valuePara);

                        return true;
                    }
                    if (prm.StorageType == StorageType.String)
                    {
                        prm.Set((string)valuePara);

                        return true;
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// Lấy ra fitting đang kết nối với ống
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="pipe"></param>
        /// <param name="conPipeNotConnect"></param>
        /// <param name="fitingCreating"></param>
        /// <returns></returns>
        public static FamilyInstance GetFittingConnected(Document doc, MEPCurve pipe, Connector conPipeNotConnect)
        {
            if (doc == null || pipe == null)
                return null;

            List<FamilyInstance> allFittings = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .OfCategory(BuiltInCategory.OST_PipeFitting)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .ToList();
            if (allFittings == null || allFittings.Count == 0)
                return null;

            List<ElementId> allElementIds = allFittings.Select(e => e.Id).ToList();

            // Create a Outline, uses a minimum and maximum XYZ point to initialize the outline.
            Outline myOutLn = CreateOutLineFromBoundingBox(pipe);
            if (myOutLn == null || myOutLn.IsEmpty)
                return null;

            // Create a BoundingBoxIntersects filter with this Outline
            BoundingBoxIntersectsFilter filter = new BoundingBoxIntersectsFilter(myOutLn);

            FilteredElementCollector collector = new FilteredElementCollector(doc, allElementIds);

            List<FamilyInstance> allFittingConnected = collector.WherePasses(filter).Cast<FamilyInstance>().ToList();

            foreach (var ele in allFittingConnected)
            {
                if (ele != null)
                {
                    if (ele.MEPModel != null)
                    {
                        foreach (Connector conNector in ele.MEPModel.ConnectorManager.Connectors)
                        {
                            if (conNector != null)
                            {
                                if (IsEqual(conNector.Origin.X, conPipeNotConnect.Origin.X)
                                    && IsEqual(conNector.Origin.Y, conPipeNotConnect.Origin.Y)
                                    && IsEqual(conNector.Origin.Z, conPipeNotConnect.Origin.Z))
                                {
                                    return ele;
                                }
                            }
                        }
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// Lấy connector ở vị trí thấp nhất và cao nhất
        /// </summary>
        /// <param name="pipe"></param>
        /// <param name="conHigher"></param>
        /// <returns></returns>
        public static Connector GetConnectorMinZ(MEPCurve pipe, out Connector conHigher)
        {
            Connector retval = null;
            conHigher = null;

            if (pipe != null)
            {
                ConnectorManager connectorManager = pipe.ConnectorManager;

                double maxZ = double.MaxValue;
                double minZ = double.MinValue;

                foreach (Connector item in connectorManager.Connectors)
                {
                    // lấy connector thấp nhất
                    if (item.Origin.Z < maxZ)
                    {
                        maxZ = item.Origin.Z;
                        retval = item;
                    }
                    // lấy connector cao nhất
                    if (item.Origin.Z > minZ)
                    {
                        minZ = item.Origin.Z;
                        conHigher = item;
                    }
                }
            }

            return retval;
        }

        /// <summary>
        /// Kiểm tra xem pipe có là ống thẳng đứng không
        /// </summary>
        /// <param name="pipeCheck"></param>
        /// <returns></returns>
        public static bool IsPipeVertical(MEPCurve pipeCheck)
        {
            if (pipeCheck == null)
                return false;

            Connector con1 = pipeCheck.ConnectorManager.Lookup(0);
            Connector con2 = pipeCheck.ConnectorManager.Lookup(1);

            if (con1 != null && con2 != null)
            {
                if (Common.IsEqual(con1.Origin.X, con2.Origin.X)
                  && Common.IsEqual(con1.Origin.Y, con2.Origin.Y))
                    return true;
            }

            return false;
        }

        /// <summary>
        /// GetPointOnVector
        /// </summary>
        /// <param name="pointInsert"></param>
        /// <param name="vectorDir"></param>
        /// <param name="dDistance"></param>
        /// <returns></returns>
        public static XYZ GetPointOnVector(XYZ pointInsert, XYZ vectorDir, double dDistance)
        {
            return (pointInsert + (vectorDir.Normalize()) * dDistance);
        }

        /// <summary>
        /// Kiểm tra góc giữa 2 pipe có hợp lệ để tạo curve
        /// </summary>
        /// <param name="pipeHor"></param>
        /// <param name="pipeVer"></param>
        /// <returns></returns>
        public static bool IsAngleValid(MEPCurve pipeHor, MEPCurve pipeVer)
        {
            if (pipeHor != null && pipeVer != null && pipeHor.Location is LocationCurve locationHor && pipeVer.Location is LocationCurve locationVer)
            {
                XYZ pointInter = GetPointIntersecNotInXYPlane(pipeVer, pipeHor);
                if (pointInter == null)
                    return false;

                Connector conHor1 = GetConnectorNearest(pointInter, pipeHor, out Connector conHor2);

                Connector conVer1 = GetConnectorNearest(pointInter, pipeVer, out Connector conVer2);

                if (conHor1 == null || conHor2 == null || conVer1 == null || conVer2 == null)
                    return false;

                XYZ vec1 = (conHor2.Origin - conHor1.Origin).Normalize();

                XYZ vec2 = (conVer2.Origin - conVer1.Origin).Normalize();

                XYZ normal = vec1.CrossProduct(vec2).Normalize();

                if (vec1 != null && vec2 != null)
                {
                    double angle = vec1.AngleOnPlaneTo(vec2, normal);
                    if (angle > Math.PI / 2 && angle < Math.PI || Common.IsEqual(angle, Math.PI / 2))
                        return true;
                }
            }

            return false;
        }

        public static Parameter GetParameterFromListedNames(Element el, params string[] args)
        {
            foreach (string str in args)
            {
                Parameter prm = el.LookupParameter(str);
                if (prm != null)
                {
                    return prm;
                }
            }

            return null;
        }

        /// <summary>
        /// Lấy chiều dài của elbow và kc từ tâm đến connector
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="pipeHorizontal"></param>
        /// <param name="vectorMoveZ"></param>
        /// <param name="lenghtElbow"></param>
        /// <param name="distanceCenterToConnector"></param>
        /// <returns></returns>
        public static bool GetLenghtElbowFitting45(Document doc, Level level, MEPCurve pipeHorizontal, XYZ vectorMoveZ
           , out double lenghtElbow, out double distanceCenterToConnector)
        {
            lenghtElbow = 0;
            distanceCenterToConnector = 0;

            if (doc == null || pipeHorizontal == null || Common.IsPipeVertical(pipeHorizontal))
                return false;

            SubTransaction subTran = new SubTransaction(doc);
            subTran.Start();
            try
            {
                Connector conHor1 = GetConnectorValid(pipeHorizontal, level);
                if (conHor1 != null && !conHor1.IsConnected)
                {
                    conHor1 = GetConnectorNearest(conHor1.Origin, pipeHorizontal, out Connector conHor2);

                    XYZ vectorMoveX = (conHor1.Origin - new XYZ(conHor2.Origin.X, conHor2.Origin.Y, conHor1.Origin.Z)).Normalize();

                    double lenghtPipeJoint = 1000 / 304.8;

                    XYZ endPoint = GetPointOnVector(conHor1.Origin, vectorMoveX, lenghtPipeJoint / Math.Sqrt(2));

                    endPoint = GetPointOnVector(endPoint, vectorMoveZ, lenghtPipeJoint / Math.Sqrt(2));

                    // Tạo ống chéo nối ống ngang và ống đứng
                    Pipe pipeConnection = Pipe.Create(doc, pipeHorizontal.GetTypeId(), pipeHorizontal.ReferenceLevel.Id, conHor1, endPoint);
                    if (pipeConnection == null)
                    {
                        subTran.RollBack();
                        subTran.Dispose();
                        return false;
                    }

                    // Tạo kết nối giữa ống ngang và ống chéo

                    FamilyInstance elbow1 = null;
                    GetConnectorClosedTo(pipeHorizontal.ConnectorManager, pipeConnection.ConnectorManager, out Connector con1, out Connector con2);
                    if (con1 != null && con2 != null && Common.IsAngleValid(pipeHorizontal, pipeConnection))
                        elbow1 = doc.Create.NewElbowFitting(con1, con2);

                    if (elbow1 == null)
                    {
                        subTran.RollBack();
                        subTran.Dispose();
                        return false;
                    }

                    Parameter paraelbow1 = GetParameterFromListedNames(elbow1, "L1");
                    if (paraelbow1 != null && paraelbow1.StorageType == StorageType.Double)
                        lenghtElbow = paraelbow1.AsDouble();

                    Connector conElbow1 = elbow1.MEPModel.ConnectorManager.Lookup(1);
                    if (conElbow1 != null)
                        distanceCenterToConnector = ((LocationPoint)elbow1.Location).Point.DistanceTo(conElbow1.Origin);
                }

                subTran.RollBack();
                subTran.Dispose();
                return true;
            }
            catch (Exception)
            {
                subTran.RollBack();
                subTran.Dispose();
                return false;
            }
        }

        public static bool IsFormSameOpen(string nameForm)
        {
            bool checkIsOpen = false;

            foreach (System.Windows.Forms.Form openedForm in Application.OpenForms)
            {
                if (openedForm.GetType().Name == nameForm)
                {
                    openedForm.WindowState = FormWindowState.Normal;
                    checkIsOpen = true;
                    break;
                }
            }

            return checkIsOpen;
        }

        /// <summary>
        /// kiểm tra xem đã set RoutingPreference cho pipe type hay chưa
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="pipe"></param>
        /// <param name="checkType">
        /// Tee: RoutingPreferenceRuleGroupType.Junctions
        /// Elbow: RoutingPreferenceRuleGroupType.Elbows
        /// </param>
        /// <returns></returns>
        public static bool IsFamilySymbolSettedForPipeType(Document doc, MEPCurve pipe, RoutingPreferenceRuleGroupType checkType)
        {
            try
            {
                if (pipe != null && pipe.IsValidObject)
                {
                    var pipeType = doc.GetElement(pipe.GetTypeId()) as PipeType;

                    RoutingPreferenceManager rpm = pipeType.RoutingPreferenceManager;

                    if (checkType == RoutingPreferenceRuleGroupType.Junctions &&
                      rpm.PreferredJunctionType != PreferredJunctionType.Tee)
                        return false;

                    int numberOfRule = rpm.GetNumberOfRules(checkType);

                    for (int i = 0; i < numberOfRule; i++)
                    {
                        RoutingPreferenceRule rule = rpm.GetRule(checkType, i);

                        if (rule.MEPPartId != null &&
                            rule.MEPPartId != ElementId.InvalidElementId)
                        {
                            return true;
                        }
                    }
                }

                return false;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Lấy symbol set RoutingPreference cho pipe type hay chưa
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="pipe"></param>
        /// <param name="checkType">
        /// Tee: RoutingPreferenceRuleGroupType.Junctions
        /// Elbow: RoutingPreferenceRuleGroupType.Elbows
        /// </param>
        /// <returns></returns>
        public static FamilySymbol GetSymbolSeted(Document doc, MEPCurve pipe, RoutingPreferenceRuleGroupType checkType)
        {
            try
            {
                if (doc != null && pipe != null && pipe.IsValidObject)
                {
                    var pipeType = doc.GetElement(pipe.GetTypeId()) as PipeType;

                    RoutingPreferenceManager rpm = pipeType.RoutingPreferenceManager;

                    if (checkType == RoutingPreferenceRuleGroupType.Junctions &&
                      rpm.PreferredJunctionType != PreferredJunctionType.Tee)
                        return null;

                    int numberOfRule = rpm.GetNumberOfRules(checkType);

                    if (numberOfRule > 0)
                    {
                        RoutingPreferenceRule rule = rpm.GetRule(checkType, numberOfRule - 1);

                        if (rule.MEPPartId != null &&
                            rule.MEPPartId != ElementId.InvalidElementId)
                        {
                            return doc.GetElement(rule.MEPPartId) as FamilySymbol;
                        }
                    }
                }

                return null;
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>
        /// Lấy ra connector gần nhất và xa nhất với 1 điểm cho trước
        /// </summary>
        /// <param name="point"></param>
        /// <param name="pipe"></param>
        /// <param name="outFarest"></param>
        /// <returns></returns>
        public static Connector GetConnectorNearest(XYZ point, ConnectorManager connectorManager)
        {
            Connector retval = null;

            if (point != null && connectorManager != null)
            {
                double max = double.MaxValue;
                double min = double.MinValue;

                foreach (Connector item in connectorManager.Connectors)
                {
                    double distance = item.Origin.DistanceTo(point);

                    // lấy connector gần nhất
                    if (distance < max)
                    {
                        max = distance;
                        retval = item;
                    }
                }
            }

            return retval;
        }

        public static XYZ GetUnBoundIntersection(Line Line1, Line Line2)
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

        private static void RotateLine(Document doc, FamilyInstance wye, Line axisLine)
        {
            Line rotateLine = Line.CreateBound(wye.MEPModel.ConnectorManager.Lookup(1).Origin, (wye.Location as LocationPoint).Point);

            XYZ vector = rotateLine.Direction.CrossProduct(axisLine.Direction);
            XYZ intersection = GetUnBoundIntersection(rotateLine, axisLine);

            double angle = rotateLine.Direction.AngleTo(axisLine.Direction);

            Line line = Line.CreateUnbound(intersection, vector);

            ElementTransformUtils.RotateElement(doc, wye.Id, line, angle);
            doc.Regenerate();
        }

        /// <summary>
        /// Get end length of the eblow to the eblow (dùng sub transaction)
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="diameter"></param>
        /// <param name="symbolEblow"></param>
        /// <returns></returns>
        public static double GetEndLengthOfEblowSubTran(Document doc, double diameter, out double distanceCenterToConnector, FamilySymbol symbolEblow = null)
        {
            double length = 0;
            distanceCenterToConnector = 0.0;
            if (symbolEblow == null)
                symbolEblow = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol))
                                                                         .Cast<FamilySymbol>()
                                                                         .FirstOrDefault(x => x.FamilyName.Equals("uPVC_Elbow_TienPhong"));
            if (symbolEblow != null)
            {
                using (SubTransaction tran = new SubTransaction(doc))
                {
                    tran.Start();
                    try
                    {
                        FamilyInstance instance = doc.Create.NewFamilyInstance(XYZ.Zero, symbolEblow, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                        Parameter paraDiameter = instance.LookupParameter("Nominal Diameter");
                        if (paraDiameter != null && !paraDiameter.IsReadOnly && paraDiameter.StorageType == StorageType.Double)
                            paraDiameter.Set(diameter);
                        else
                        {
                            paraDiameter = instance.LookupParameter("ENG_Nominal Diameter");
                            if (paraDiameter != null && !paraDiameter.IsReadOnly && paraDiameter.StorageType == StorageType.Double)
                                paraDiameter.Set(diameter);
                        }
                        doc.Regenerate();

                        Connector connector = instance.MEPModel?.ConnectorManager.Lookup(1);
                        if (connector != null)
                            distanceCenterToConnector = connector.Origin.DistanceTo(XYZ.Zero);

                        Parameter paraL1 = GetParameterFromListedNames(instance, "L1");
                        if (paraL1 != null && paraL1.StorageType == StorageType.Double)
                            return paraL1.AsDouble();

                        if (paraL1 == null)
                            return length = 10 / 304.8;
                    }
                    catch (Exception) { }
                    finally
                    {
                        tran.RollBack();
                    }
                }
            }

            return length;
        }

        /// <summary>
        /// Kết nối ống thẳng đứng với ống ngang bằng kết nối dạng góc 45 độ (Elbow 45)
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="pipe1"></param>
        /// <param name="pipe2"></param>
        /// <param name="isHubMode"></param>
        public static bool ConnectPipeVerticalElbow45(Document doc, Level level, MEPCurve pipe1, MEPCurve pipe2, bool isHubMode)
        {
            if (pipe1 == null || pipe2 == null)
                return false;

            try
            {
                if (pipe1 != null && pipe2 != null && pipe1.Location is LocationCurve && pipe2.Location is LocationCurve)
                {
                    // Di chuyển 2 pipe về cùng tâm
                    MovePipeToCenter(doc, pipe1, pipe2);

                    bool isSucces = DetermindPipeVertcalAndHorizontal(pipe1, pipe2, out MEPCurve pipeVertical, out MEPCurve pipeHorizontal);

                    GetConnectorClosedTo(pipeVertical.ConnectorManager, pipeHorizontal.ConnectorManager, out Connector conSt, out Connector conEnd);
                    if (conSt == null || conSt.IsConnected || conEnd == null || conEnd.IsConnected)
                        return false;

                    if (isSucces)
                    {
                        // tìm điểm giao nhau
                        XYZ pointIntersec = GetPointIntersecNotInXYPlane(pipeVertical, pipeHorizontal);

                        if (pointIntersec == null)
                            return false;

                        // đưa 2 ống về cùng kích thước
                        if (pipeHorizontal.Diameter != pipeVertical.Diameter)
                            SetValueParameterByBuiltIn(pipeHorizontal, BuiltInParameter.RBS_PIPE_DIAMETER_PARAM, pipeVertical.Diameter);

                        double slope = Math.Round((double)GetValueParameterByBuilt(pipeHorizontal, BuiltInParameter.RBS_PIPE_SLOPE), 5);

                        Connector conHor1 = GetConnectorNearest(pointIntersec, pipeHorizontal, out Connector conHor2);

                        XYZ vectorMoveX = (conHor2.Origin - new XYZ(pointIntersec.X, pointIntersec.Y, conHor2.Origin.Z)).Normalize();

                        // ống thẳng đứng
                        Connector conVert1 = GetConnectorNearest(pointIntersec, pipeVertical, out Connector conVert2);

                        XYZ vectoMoveZ = (conVert2.Origin - pointIntersec).Normalize();

                        if (!GetLenghtElbowFitting45(doc, level, pipeHorizontal, vectoMoveZ, out double lengthElbow1, out double distanceCenterToConnector1))
                            return false;

                        FamilySymbol typeElbow = GetSymbolSeted(doc, pipeHorizontal, RoutingPreferenceRuleGroupType.Elbows);

                        double lengthElbow2 = GetEndLengthOfEblowSubTran(doc, pipeHorizontal.Diameter, out double distanceCenterToConnector2, typeElbow);
                        double lenghtPipeJoint = lengthElbow1 + lengthElbow2;

                        if (Common.IsEqual(lenghtPipeJoint, 0))
                            lenghtPipeJoint = 10 / 304.8;

                        XYZ newEndPointElbow1 = GetPointOnVector(conVert1.Origin, vectoMoveZ.Negate(), lenghtPipeJoint);
                        newEndPointElbow1 = GetPointOnVector(newEndPointElbow1, vectorMoveX, lenghtPipeJoint);

                        // Tạo ống chéo nối ống ngang và ống đứng
                        Pipe pipeConnection = Pipe.Create(doc, pipeVertical.GetTypeId(), pipeVertical.ReferenceLevel.Id, conVert1, newEndPointElbow1);
                        if (pipeConnection == null)
                            return false;

                        // Tạo kết nối giữa ống đứng và ống chéo
                        GetConnectorClosedTo(pipeVertical.ConnectorManager, pipeConnection.ConnectorManager, out Connector con1, out Connector con2);

                        FamilyInstance elbowFitting1 = null;

                        if (con1 != null && con2 != null && IsAngleValid(pipeVertical, pipeConnection))
                            elbowFitting1 = doc.Create.NewElbowFitting(con1, con2);

                        if (elbowFitting1 == null)
                            return false;

                        doc.Delete(pipeConnection.Id);

                        Connector conElbow1 = GetConnectorNotConnnected(elbowFitting1.MEPModel.ConnectorManager);

                        if (conElbow1 != null)
                        {
                            XYZ vectorCheo = (conElbow1.Origin - ((LocationPoint)elbowFitting1.Location).Point).Normalize();

                            XYZ poinEndPipeConnect = Common.GetPointOnVector(conElbow1.Origin, vectorCheo, lenghtPipeJoint);

                            Pipe pipeConnectionNew2 = Pipe.Create(doc, pipeHorizontal.GetTypeId(), pipeHorizontal.ReferenceLevel.Id, conElbow1, poinEndPipeConnect);

                            if (pipeConnectionNew2 != null)
                            {
                                Connector conPipeConnect2 = GetConnectorNotConnnected(pipeConnectionNew2.ConnectorManager);
                                if (conPipeConnect2 == null)
                                    return false;

                                XYZ poinStartPipeHorizontal2 = Common.GetPointOnVector(conPipeConnect2.Origin, vectorCheo, distanceCenterToConnector1);

                                XYZ poinEndPipeHorizontal2 = Common.GetPointOnVector(poinStartPipeHorizontal2, vectorMoveX, 1000 / 304.8);
                                double distanceLenghtPipe = poinEndPipeHorizontal2.DistanceTo(poinStartPipeHorizontal2);
                                poinEndPipeHorizontal2 = Common.GetPointOnVector(poinEndPipeHorizontal2, vectoMoveZ.Negate(), distanceLenghtPipe * slope);

                                XYZ directionpipeHorizoltal2New2 = (poinEndPipeHorizontal2 - poinStartPipeHorizontal2).Normalize();

                                poinStartPipeHorizontal2 = Common.GetPointOnVector(poinStartPipeHorizontal2, directionpipeHorizoltal2New2, distanceCenterToConnector1);

                                Pipe pipeHorizoltal2New2 = Pipe.Create(doc, pipeConnectionNew2.MEPSystem.GetTypeId(), pipeConnectionNew2.GetTypeId(), pipeConnectionNew2.ReferenceLevel.Id, poinStartPipeHorizontal2, poinEndPipeHorizontal2);

                                if (pipeHorizoltal2New2 != null)
                                {
                                    SetValueParameterByBuiltIn(pipeHorizoltal2New2, BuiltInParameter.RBS_PIPE_DIAMETER_PARAM, pipeConnectionNew2.Diameter);

                                    GetConnectorClosedTo(pipeConnectionNew2.ConnectorManager, pipeHorizoltal2New2.ConnectorManager, out con1, out con2);

                                    FamilyInstance elbowFitting2 = null;

                                    if (con1 != null && con2 != null && Common.IsAngleValid(pipeHorizoltal2New2, pipeConnectionNew2))
                                        elbowFitting2 = doc.Create.NewElbowFitting(con1, con2);

                                    if (elbowFitting2 == null)
                                        return false;

                                    doc.Delete(pipeHorizoltal2New2.Id);

                                    // Kéo dài ống thẳng đứng  tới ống nằm ngang
                                    Connector conElbow2 = GetConnectorNotConnnected(elbowFitting2.MEPModel.ConnectorManager);
                                    if (conElbow2 == null)
                                        return false;

                                    conHor1 = GetConnectorNearest(pointIntersec, pipeHorizontal, out conHor2);

                                    XYZ vectorMoveZ1 = (conHor1.Origin.Z > conHor2.Origin.Z || Common.IsEqual(conHor1.Origin.Z, conHor2.Origin.Z)) ? XYZ.BasisZ : XYZ.BasisZ.Negate();

                                    XYZ newPoint2 = new XYZ(conElbow2.Origin.X, conElbow2.Origin.Y, conHor2.Origin.Z);
                                    double distance = newPoint2.DistanceTo(conHor2.Origin);

                                    newPoint2 = Common.GetPointOnVector(newPoint2, vectorMoveZ1, slope * distance);

                                    double height = Math.Abs(conElbow2.Origin.Z - newPoint2.Z);

                                    XYZ vectran = (XYZ.BasisZ.Negate()) * height;

                                    if (conHor1.Origin.Z > conElbow2.Origin.Z)
                                        ElementTransformUtils.MoveElement(doc, elbowFitting2.Id, vectran.Negate());
                                    else
                                        ElementTransformUtils.MoveElement(doc, elbowFitting2.Id, vectran);

                                    ResetLocation(pipeHorizontal, conHor2.Origin, newPoint2);

                                    GetConnectorClosedTo(pipeHorizontal.ConnectorManager, elbowFitting2.MEPModel.ConnectorManager, out con1, out con2);
                                    if (con1 != null && con2 != null)
                                        con1.ConnectTo(con2);

                                    if (IsEqual(slope, 0))
                                    {
                                        Connector conStPipe = pipeHorizontal.ConnectorManager.Lookup(0);
                                        Connector conEndPipe = pipeHorizontal.ConnectorManager.Lookup(1);

                                        if (conStPipe != null && conEndPipe != null)
                                        {
                                            if (!conStPipe.IsConnected || !conEndPipe.IsConnected)
                                            {
                                                Connector conNotConnected = (!conStPipe.IsConnected) ? conStPipe : conEndPipe;

                                                FamilyInstance elbowFittingConnected = GetFittingConnected(doc, pipeHorizontal, conNotConnected, elbowFitting2.Id);
                                                if (elbowFittingConnected != null)
                                                {
                                                    GetConnectorClosedTo(pipeHorizontal.ConnectorManager, elbowFittingConnected.MEPModel.ConnectorManager, out con1, out con2);
                                                    if (con1 != null && con2 != null && !con1.IsConnected && !con2.IsConnected)
                                                        con1.ConnectTo(con2);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    return true;
                }

                return false;
            }
            catch (Exception)
            {
                //IO.ShowError(ex.Message, "Error");
                return false;
            }
        }

        /// <summary>
        /// Lấy ra fitting đang kết nối với ống
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="pipe"></param>
        /// <param name="conPipeNotConnect"></param>
        /// <param name="fitingCreating"></param>
        /// <returns></returns>
        public static FamilyInstance GetFittingConnected(Document doc, MEPCurve pipe, Connector conPipeNotConnect, ElementId fitingCreating)
        {
            if (doc == null || pipe == null || fitingCreating == ElementId.InvalidElementId)
                return null;

            List<FamilyInstance> allFittings = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .OfCategory(BuiltInCategory.OST_PipeFitting)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .ToList();
            if (allFittings == null || allFittings.Count == 0)
                return null;

            List<ElementId> allElementIds = allFittings.Select(e => e.Id).ToList();

            // Create a Outline, uses a minimum and maximum XYZ point to initialize the outline.
            Outline myOutLn = CreateOutLineFromBoundingBox(pipe);
            if (myOutLn == null || myOutLn.IsEmpty)
                return null;

            // Create a BoundingBoxIntersects filter with this Outline
            BoundingBoxIntersectsFilter filter = new BoundingBoxIntersectsFilter(myOutLn);

            FilteredElementCollector collector = new FilteredElementCollector(doc, allElementIds);

            List<FamilyInstance> allFittingConnected = collector.WherePasses(filter).Cast<FamilyInstance>().ToList();

            foreach (var ele in allFittingConnected)
            {
                if (ele != null && ele.Id != fitingCreating)
                {
                    if (ele.MEPModel != null)
                    {
                        foreach (Connector conNector in ele.MEPModel.ConnectorManager.Connectors)
                        {
                            if (conNector != null)
                            {
                                if (IsEqual(conNector.Origin.X, conPipeNotConnect.Origin.X)
                                    && IsEqual(conNector.Origin.Y, conPipeNotConnect.Origin.Y)
                                    && IsEqual(conNector.Origin.Z, conPipeNotConnect.Origin.Z))
                                {
                                    return ele;
                                }
                            }
                        }
                    }
                }
            }

            return null;
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

        /// <summary>
        /// Lấy thông tin của dộ dốc
        /// </summary>
        /// <param name="pipe">
        ///       slope = 0 => Không dốc  </param>
        ///       slope = 1 => Dốc lên  </param>
        ///        slope = 2 => Dốc xuống  </param>
        /// <returns></returns>
        public static int GetSlope(MEPCurve pipe)
        {
            int option = 0;
            if (pipe != null && pipe.Location is LocationCurve location)
            {
                XYZ stPoint = location.Curve.GetEndPoint(0);
                XYZ endpoint = location.Curve.GetEndPoint(1);

                if (stPoint != null && endpoint != null)
                {
                    if (IsEqual(stPoint.Z, endpoint.Z))
                        return 0; // Slope Off

                    if (endpoint.Z > stPoint.Z)
                        return 1; //Slope up

                    if (endpoint.Z < stPoint.Z)
                        return 2; // Slope down
                }
            }

            return option;
        }

        /// <summary>
        /// Reset vị trí của pipe
        /// </summary>
        /// <param name="pipe"></param>
        /// <param name="point1"></param>
        /// <param name="point2"></param>
        /// <returns></returns>
        public static bool ResetLocation(MEPCurve pipe, XYZ point1, XYZ point2)
        {
            if (pipe != null && point1 != null && point2 != null && point1.IsAlmostEqualTo(point2) == false)
            {
                int optionSlope = GetSlope(pipe);

                XYZ stPoint = (point1.Z > point2.Z) ? point1 : point2;
                XYZ endPoint = (point1.Z < point2.Z) ? point1 : point2;

                Line lcNew = Line.CreateBound(point1, point2);

                if (optionSlope == 0)//Slope up
                    lcNew = Line.CreateBound(point1, point2);//Slope off
                else if (optionSlope == 1)
                    lcNew = Line.CreateBound(endPoint, stPoint);//Slope up
                else
                    lcNew = Line.CreateBound(stPoint, endPoint);//Slope down

                if (lcNew != null)
                    ((LocationCurve)pipe.Location).Curve = lcNew;
            }

            return false;
        }

        /// <summary>
        /// Lấy connector chưa được kết nối
        /// </summary>
        /// <param name="connectorManager"></param>
        /// <returns></returns>
        public static Connector GetConnectorNotConnnected(ConnectorManager connectorManager)
        {
            if (connectorManager != null)
            {
                foreach (Connector con in connectorManager.Connectors)
                {
                    if (!con.IsConnected)
                        return con;
                }
            }

            return null;
        }

        public static Solid CreateCylindricalVolume(XYZ point, double height, double radius, bool bUp)
        {
            // build cylindrical shape around endpoint
            List<CurveLoop> curveloops = new List<CurveLoop>();
            CurveLoop circle = new CurveLoop();

            // For solid geometry creation, two curves are necessary, even for closed
            // cyclic shapes like circles

            Arc arc1 = Arc.Create(point, radius, 0, Math.PI, XYZ.BasisX, XYZ.BasisY);
            Arc arc2 = Arc.Create(point, radius, Math.PI, 2 * Math.PI, XYZ.BasisX, XYZ.BasisY);

            circle.Append(arc1);
            circle.Append(arc2);
            curveloops.Add(circle);

            Solid createdCylinder = GeometryCreationUtilities.CreateExtrusionGeometry(curveloops, XYZ.BasisZ * (bUp ? 1 : -1), height);

            return createdCylinder;
        }

        public static List<Connector> GetConnectionNearest(Element mep1, Element mep2)
        {
            try
            {
                Curve curve1 = GetCurve(mep1);
                Curve curve2 = GetCurve(mep2);

                XYZ start1 = curve1.GetEndPoint(0);
                XYZ end1 = curve1.GetEndPoint(1);

                XYZ start2 = curve2.GetEndPoint(0);
                XYZ end2 = curve2.GetEndPoint(1);

                Connector c1 = null;
                Connector c2 = null;

                if (start1.DistanceTo(start2) < end1.DistanceTo(start2))
                {
                    c1 = GetConnectorClosestTo(mep1, start1);
                }
                else
                {
                    c1 = GetConnectorClosestTo(mep1, end1);
                }

                if (start2.DistanceTo(start1) < end2.DistanceTo(start1))
                {
                    c2 = GetConnectorClosestTo(mep2, start2);
                }
                else
                {
                    c2 = GetConnectorClosestTo(mep2, end2);
                }

                if (c1 == null || c2 == null)
                    return null;

                List<Connector> connectors = new List<Connector>() { c1, c2 };

                return connectors;
            }
            catch (System.Exception ex)
            {
                return null;
            }
        }

        public static Curve GetCurve(this Element e)
        {
            Debug.Assert(null != e.Location,
              "expected an element with a valid Location");

            LocationCurve lc = e.Location as LocationCurve;

            Debug.Assert(null != lc,
              "expected an element with a valid LocationCurve");

            return lc.Curve;
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

        public static bool IsVertical(XYZ v)
        {
            return IsZero(v.X) && IsZero(v.Y);
        }

        public static bool IsZero(double a)
        {
            return IsZero(a, _eps);
        }

        /// <summary>
        /// Return the given element's connector manager,
        /// using either the family instance MEPModel or
        /// directly from the MEPCurve connector manager
        /// for ducts and pipes.
        /// </summary>
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

        public static bool IsLessThan(double first, double second)
        {
            if (!IsEqual(first, second, 10e-5))
                return first < second;

            return false;
        }

        public static bool IsGreaterThanOrEqual(double first, double second)
        {
            return IsEqual(first, second) || first > second;
        }

        public static bool IsLessThanOrEqual(double first, double second)
        {
            return IsEqual(first, second) || first < second;
        }

        public static XYZ ConvertStringToXYZ(string value)
        {
            XYZ retVal = null;
            if (value?.Length > 0)
            {
                value = value.Remove(0, 1);
                value = value.Remove(value.Length - 1, 1);

                string[] coors = value.Split(',');
                if (coors.Count() == 3)
                {
                    double X = double.Parse(coors[0]);
                    double Y = double.Parse(coors[1]);
                    double Z = double.Parse(coors[2]);
                    retVal = new XYZ(X, Y, Z);
                }
            }

            return retVal;
        }

        public enum Position
        {
            Center,
            Left,
            Right
        }
    }

    public class MyUnit
    {
        public static double da(double degrees)
        {
            return degrees / 180.0 * Math.PI;
        }

        public static double rd(double radians)
        {
            return radians * 180.0 / Math.PI;
        }
    }

    public class ObjectItem
    {
        private string _name;
        private ElementId _objectId = ElementId.InvalidElementId;

        private string Guid = null;

        public ObjectItem(string name, ElementId objectId)
        {
            _name = name;
            _objectId = objectId;
        }

        public ObjectItem(string name, string guid)
        {
            _name = name;
            Guid = guid;
        }

        public string Name
        {
            set { _name = value; }
        }

        public ElementId ObjectId
        {
            get { return _objectId; }
        }

        public override string ToString()
        {
            return _name;
        }
    }
}