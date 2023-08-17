using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Visual;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;
using TotalMEPProject.SelectionFilters;

namespace TotalMEPProject.Ultis
{
    public static class Common
    {
        private const double _inch = 1.0 / 12.0;
        public static double _sixteenth = _inch / 16.0;
        public static double mmToFT = 0.0032808399;
        private const double _eps = 1.0e-9;

        /// <summary>
        /// Get Information Connector
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="symbolTee"></param>
        /// <param name="idConSt"></param>
        /// <param name="idConEnd"></param>
        /// <param name="idConTee"></param>
        /// <returns></returns>
        public static bool GetInforConnector(Document doc, FamilySymbol symbolTee, out int idConSt, out int idConEnd, out int idConTee)
        {
            idConSt = int.MaxValue;
            idConEnd = int.MaxValue;
            idConTee = int.MaxValue;

            if (symbolTee != null)
            {
                using (SubTransaction tran = new SubTransaction(doc))
                {
                    tran.Start();
                    try
                    {
                        FamilyInstance fitting = doc.Create.NewFamilyInstance(XYZ.Zero, symbolTee, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                        GetInformationConectorWye(fitting, null, out Connector conSt, out Connector conEnd, out Connector conTee);

                        if (conSt != null)
                            idConSt = conSt.Id;
                        if (conEnd != null)
                            idConEnd = conEnd.Id;
                        if (conTee != null)
                            idConTee = conTee.Id;

                        if (idConSt != int.MaxValue && idConEnd != int.MaxValue && idConTee != int.MaxValue)
                            return true;
                    }
                    catch (Exception) { }
                    finally
                    {
                        tran.RollBack();
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// Get information of connector tee
        /// </summary>
        /// <param name="fitting"></param>
        /// <param name="vector"></param>
        /// <param name="main1"></param>
        /// <param name="main2"></param>
        /// <param name="tee"></param>
        public static void GetInformationConectorWye(FamilyInstance fitting, XYZ vector, out Connector main1, out Connector main2, out Connector tee)
        {
            main1 = null;
            main2 = null;
            tee = null;
            if (fitting != null)
            {
                //Get fitting info

                GetConnectorMain(fitting, vector, out main1, out main2);

                foreach (Connector c in fitting.MEPModel.ConnectorManager.Connectors)
                {
                    if (c.Id != main1.Id && c.Id != main2.Id)
                    {
                        tee = c;
                        break;
                    }
                }
            }
        }

        /// <summary>
        /// Get connector main
        /// </summary>
        /// <param name="fitting"></param>
        /// <param name="vector"></param>
        /// <param name="mainConnect1"></param>
        /// <param name="mainConnect2"></param>
        public static void GetConnectorMain(FamilyInstance fitting, XYZ vector, out Connector mainConnect1, out Connector mainConnect2)
        {
            mainConnect1 = null;
            mainConnect2 = null;

            MechanicalFitting mechanicalFitting = fitting.MEPModel as MechanicalFitting;
            if (mechanicalFitting != null && mechanicalFitting.PartType == PartType.Tee && vector == null && fitting.MEPModel.ConnectorManager.Connectors.Size == 3)
            {
                //Main : hướng connector của 2 connector fai song song voi nhau (nguoc chieu nhau)

                foreach (Connector c1 in fitting.MEPModel.ConnectorManager.Connectors)
                {
                    foreach (Connector c2 in fitting.MEPModel.ConnectorManager.Connectors)
                    {
                        if (c1.Id == c2.Id)
                        {
                            continue;
                        }
                        else
                        {
                            var z1 = c1.CoordinateSystem.BasisZ;
                            var z2 = c2.CoordinateSystem.BasisZ;

                            if (IsParallel(z1, z2, 0.0001) == true)
                            {
                                mainConnect1 = c1;
                                mainConnect2 = c2;
                                break;
                            }
                        }
                    }

                    if (mainConnect1 != null && mainConnect2 != null)
                        break;
                }
            }
            else
            {
                foreach (Connector con in fitting.MEPModel.ConnectorManager.Connectors)
                {
                    if (vector != null)
                    {
                        if (IsParallel(vector, con.CoordinateSystem.BasisZ, 0.0001) == false)
                        {
                            continue;
                        }
                    }

                    if (mainConnect1 == null)
                        mainConnect1 = con;
                    else
                    {
                        mainConnect2 = con;
                        break;
                    }
                }
            }

            if (mainConnect1 != null && mainConnect2 != null)
            {
                //Connect nao gan location of fitting thi do la 1

                var p = (fitting.Location as LocationPoint).Point;
                if (mainConnect1.Origin.DistanceTo(p) > mainConnect2.Origin.DistanceTo(p))
                {
                    Connector temp = mainConnect1;
                    mainConnect1 = mainConnect2;

                    mainConnect2 = temp;
                }
            }
        }

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

        public static List<Curve> CreateOffsetCurve(Curve curveOrigin, double distance, bool allSide)
        {
            if (curveOrigin is Line)
            {
                return CalculateOffsetLine(curveOrigin as Line, distance, allSide);
            }
            else if (curveOrigin is Arc)
            {
                return CalculateOffsetArc(curveOrigin, distance);
            }
            return null;
        }

        private static XYZ MiddlePoint(Curve curve)
        {
            double d1 = curve.GetEndParameter(0);
            double d2 = curve.GetEndParameter(1);
            return curve.Evaluate(d1 + ((d2 - d1) / 2.0), false);
        }

        public static List<Curve> CalculateOffsetArc(Curve curveOri, double distance)
        {
            List<Curve> offsets = new List<Curve>();

            Arc arc = curveOri as Arc;
            //Middle point in arc

            XYZ midPoint = MiddlePoint(arc);

            XYZ p1 = curveOri.GetEndPoint(0);
            XYZ p2 = curveOri.GetEndPoint(1);

            //Radius new arc

            double radius1 = arc.Radius - distance;
            double radius2 = arc.Radius + distance;

            List<double> list_radius = new List<double>();
            list_radius.Add(radius1);
            list_radius.Add(radius2);
            for (int i = 0; i < list_radius.Count; i++)
            {
                double radius = list_radius[i];
                XYZ normal = (arc.Center - midPoint);

                Line line = Line.CreateUnbound(arc.Center, normal);
                IntersectionResultArray resultArray = null;
                SetComparisonResult result = arc.Intersect(line, out resultArray);
                if (result != SetComparisonResult.Overlap)
                {
                    return null;
                }
                IntersectionResult inter = resultArray.get_Item(0);
                XYZ pOnArcOri = inter.XYZPoint;

                //Tạo Arc
                Arc arcNew = Arc.Create(arc.Center, radius, 0, 2 * Math.PI, arc.XDirection, arc.YDirection);
                XYZ normal0 = (midPoint - arc.Center);//(arc.Center - midPoint);
                Line lineFromCenterToInters = Line.CreateUnbound(arc.Center, normal0);
                XYZ normal1 = (p1 - arc.Center);//(arc.Center - p1);
                Line lineFromCenterToP1 = Line.CreateUnbound(arc.Center, normal1);
                XYZ normal2 = (p2 - arc.Center);// (arc.Center - p2);
                Line lineFromCenterToP2 = Line.CreateUnbound(arc.Center, normal2);

                //Find intersection : lineFromCenterToP1
                result = arcNew.Intersect(lineFromCenterToP1, out resultArray);
                if (result != SetComparisonResult.Overlap)
                {
                    return null;
                }
                //---------------------NOTE: NẾU NORMAL = POINT - CENTER POINT ->LẤY POINT INTERSECT TẠI 0
                //---------------------NOTE: NẾU NORMAL = CENTER POINT - POINT ->LẤY POINT INTERSECT TẠI 1

                XYZ p11 = resultArray.get_Item(0).XYZPoint;

                //Find intersection : lineFromCenterToP1
                result = arcNew.Intersect(lineFromCenterToP2, out resultArray);
                if (result != SetComparisonResult.Overlap)
                {
                    return null;
                }
                XYZ p22 = resultArray.get_Item(0).XYZPoint;

                //Find intersection : lineFromCenterToInters
                result = arcNew.Intersect(lineFromCenterToInters, out resultArray);
                if (result != SetComparisonResult.Overlap)
                {
                    return null;
                }
                XYZ pOnArc = resultArray.get_Item(0).XYZPoint;

                var arc_new = Arc.Create(p11, p22, pOnArc);
                offsets.Add(arc_new);

                //Common.CreateModelArc(arc_new);
            }

            return offsets;
        }

        public static void SetOffset(Element element, double offset)
        {
#if RV_2016 || RV_2017 || RV_2018 || RV_2019
            element.LookupParameter("Offset").Set(offset * Common.mmToFT);

#else
            element.LookupParameter("Middle Elevation").Set(offset * Common.mmToFT);

#endif
        }

        public static bool IsTap(MEPCurve mep)
        {
            if (mep as Pipe != null)
            {
                var pipe = mep as Pipe;

                var pipeType = pipe.PipeType as PipeType;

                if (pipeType.RoutingPreferenceManager.PreferredJunctionType == PreferredJunctionType.Tap)
                    return true;
            }
            else if (mep as Duct != null)
            {
                var duct = mep as Duct;
                var ductType = duct.DuctType as DuctType;

                if (ductType.RoutingPreferenceManager.PreferredJunctionType == PreferredJunctionType.Tap)
                    return true;
            }
            //             else if (mep as CableTray != null)
            //             {
            //                 var duct = mep as CableTray;
            //                 var ductType = duct. as DuctType;
            //
            //                 if (ductType.RoutingPreferenceManager.PreferredJunctionType == PreferredJunctionType.Tap)
            //                     return true;
            //             }
            //             else if (mep as Duct != null)
            //             {
            //                 var duct = mep as Duct;
            //                 var ductType = duct.DuctType as DuctType;
            //
            //                 if (ductType.RoutingPreferenceManager.PreferredJunctionType == PreferredJunctionType.Tap)
            //                     return true;
            //             }

            return false;
        }

        public static LinePatternElement CreateLineParttern(Document doc, string name)
        {
            try
            {
                FilteredElementCollector fec = new FilteredElementCollector(doc).OfClass(typeof(LinePatternElement));
                var list = fec.ToElements();

                foreach (LinePatternElement linePatternElement in list)
                {
                    if (linePatternElement.Name == name)
                        return linePatternElement;
                }

                LinePattern LinePattern = new LinePattern(name);

                List<LinePatternSegment> segments = new List<LinePatternSegment>();
                var linePatternSegment = new LinePatternSegment(LinePatternSegmentType.Dash, 3 * Common.mmToFT);
                segments.Add(linePatternSegment);
                linePatternSegment = new LinePatternSegment(LinePatternSegmentType.Space, 3 * Common.mmToFT);
                segments.Add(linePatternSegment);

                LinePattern.SetSegments(segments);

                return LinePatternElement.Create(Global.UIDoc.Document, LinePattern);
            }
            catch (System.Exception ex)
            {
                return null;
            }
        }

        public static List<Curve> CalculateOffsetLine(Line curveOrigin, double distance, bool allSide)
        {
            List<Curve> offsets = new List<Curve>();

            List<double> offset_values = new List<double>();
            if (allSide == true)
            {
                offset_values.Add(distance);
                offset_values.Add(-distance);
            }
            else
                offset_values.Add(distance);

            foreach (double d in offset_values)
            {
                Line line = Line.CreateBound(curveOrigin.GetEndPoint(0), curveOrigin.GetEndPoint(1));
                XYZ lineDirection = line.Direction;
                XYZ normal = XYZ.BasisZ.CrossProduct(lineDirection).Normalize();

                XYZ translation = normal.Multiply(d);

                XYZ startPointOffset = line.GetEndPoint(0).Add(translation);
                XYZ endPointOffset = line.GetEndPoint(1).Add(translation);
                XYZ midPoint = new XYZ((startPointOffset.X + endPointOffset.X) / 2, (startPointOffset.Y + endPointOffset.Y) / 2, (startPointOffset.Z + endPointOffset.Z) / 2);

                var curveOffset = Line.CreateBound(startPointOffset, endPointOffset);
                offsets.Add(curveOffset);
            }

            //Common.CreateModelLine(startPointOffset, endPointOffset);
            //Common.CreateModelLine(startPointOffset_2, endPointOffset_2);

            return offsets;
        }

        public static XYZ GetPipeDirection(MEPCurve mepCurve)
        {
            Curve c = mepCurve.GetCurve();
            XYZ dir = c.GetEndPoint(1) - c.GetEndPoint(0);
            dir = dir.Normalize();
            return dir;
        }

        public static int Compare(double a, double b)
        {
            return IsEqual(a, b) ? 0 : (a < b ? -1 : 1);
        }

        public static int Compare(XYZ p, XYZ q)
        {
            int d = Compare(p.X, q.X);

            if (0 == d)
            {
                d = Compare(p.Y, q.Y);

                if (0 == d)
                {
                    d = Compare(p.Z, q.Z);
                }
            }
            return d;
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

        public static void RotateLineC2(Document doc, FamilyInstance wye, Line axisLine)
        {
            var lst = Common.ToList(wye.MEPModel.ConnectorManager.Connectors);

            Connector connector2 = lst[0];
            Connector connector3 = lst[1];

            Line rotateLine = Line.CreateBound(connector2.Origin, connector3.Origin);

            if (IsParallel(axisLine.Direction, rotateLine.Direction))
                return;

            XYZ vector = rotateLine.Direction.CrossProduct(axisLine.Direction);
            XYZ intersection = GetUnBoundIntersection(rotateLine, axisLine);

            double angle = rotateLine.Direction.AngleTo(axisLine.Direction);

            Line line = Line.CreateUnbound(intersection, vector);

            ElementTransformUtils.RotateElement(doc, wye.Id, line, angle);
            doc.Regenerate();
        }

        public static void RotateLine(Document doc, FamilyInstance wye, Line axisLine)
        {
            Line rotateLine = Line.CreateBound(wye.MEPModel.ConnectorManager.Lookup(1).Origin, (wye.Location as LocationPoint).Point);

            XYZ vector = rotateLine.Direction.CrossProduct(axisLine.Direction);
            XYZ intersection = GetUnBoundIntersection(rotateLine, axisLine);

            double angle = rotateLine.Direction.AngleTo(axisLine.Direction);

            Line line = Line.CreateUnbound(intersection, vector);

            ElementTransformUtils.RotateElement(doc, wye.Id, line, angle);
            doc.Regenerate();
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
        public static Connector GetConnectorValid(MEPCurve pipe, double offset)
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
                    if (levelPipe.Elevation < offset)
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

        public static string FeetToMmString(double a)
        {
            return (a / Common.mmToFT).ToString("0.##");
        }

        public static List<Solid> GetSolid(Pipe pipe)
        {
            var options = new Options();
            options.ComputeReferences = false;

            var goes = pipe.get_Geometry(options);

            List<Solid> list = new List<Solid>();
            Common.GetSolid(goes, list, false);

            if (list.Count == 0)
                return null;

            return list;
        }

        public static void GetSolid(GeometryObject geObject, List<Solid> listSolid, bool getSymbol = false)
        {
            if (geObject is Solid)
            {
                Solid solid = geObject as Solid;
                listSolid.Add(geObject as Solid);
            }
            if (geObject as GeometryElement != null)
            {
                GeometryElement geo = geObject as GeometryElement;
                IEnumerator<GeometryObject> Objects = geo.GetEnumerator();
                while (Objects.MoveNext())
                {
                    GeometryObject geObject1 = Objects.Current;
                    GetSolid(geObject1, listSolid, getSymbol);
                }
            }
            if (geObject as GeometryInstance != null)
            {
                GeometryInstance geometryInstance = geObject as GeometryInstance;
                GeometryElement geo = null;
                if (getSymbol == true)
                    geo = geometryInstance.GetSymbolGeometry(); //_NOTE: GetSymbolGeometry de lay duoc face co reference = Mullion hien tai
                else
                    geo = geometryInstance.GetInstanceGeometry();
                IEnumerator<GeometryObject> Objects = geo.GetEnumerator();
                while (Objects.MoveNext())
                {
                    GeometryObject geObject1 = Objects.Current;
                    GetSolid(geObject1, listSolid, getSymbol);
                }
            }
        }

        public static bool IsParallel(MEPCurve p1, MEPCurve p2)
        {
            Line c1 = p1.GetCurve() as Line;
            Line c2 = p2.GetCurve() as Line;
            return Math.Sin(c1.Direction.AngleTo(
              c2.Direction)) < 0.01;
        }

        public static ConnectorProfileType GetShape(MEPCurve mep)
        {
            ConnectorProfileType ductShape
              = ConnectorProfileType.Invalid;

            foreach (Connector c
              in mep.ConnectorManager.Connectors)
            {
                if (c.ConnectorType == ConnectorType.End)
                {
                    ductShape = c.Shape;
                    break;
                }
            }
            return ductShape;
        }

        public static bool IsParallel(XYZ p, XYZ q, double tolerance)
        {
            if (p.CrossProduct(q).IsZeroLength() == true)
                return true;

            var l = p.CrossProduct(q).GetLength();
            if (IsZero(l, tolerance))
                return true;

            return false;
        }

        public static void GetInfo(FamilyInstance fitting, XYZ vector, out Connector main1, out Connector main2, out Connector tee)
        {
            //Get fitting info
            main1 = null;
            main2 = null;
            tee = null;
            Rotate45Utils.mc(fitting, vector, out main1, out main2);

            foreach (Connector c in fitting.MEPModel.ConnectorManager.Connectors)
            {
                if (c.Id != main1.Id && c.Id != main2.Id)
                {
                    tee = c;
                    break;
                }
            }
        }

        public static Element Clone(Element element)
        {
            //Create new pipe
            var newPlace = new XYZ(0, 0, 0);
            var elemIds = ElementTransformUtils.CopyElement(
              Global.UIDoc.Document, element.Id, newPlace);

            var clone = Global.UIDoc.Document.GetElement(elemIds.ToList()[0]);

            return clone;
        }

        public static void DisconnectFrom(Pipe pipe)
        {
            if (pipe != null)
            {
                Connector con1 = pipe.ConnectorManager.Lookup(0);
                Connector con2 = pipe.ConnectorManager.Lookup(1);

                if (con1 != null && con1.IsConnected)
                {
                    foreach (Connector item in con1.AllRefs)
                    {
                        if (item != null && item.IsConnectedTo(con1))
                        {
                            con1.DisconnectFrom(item);
                        }
                    }
                }

                if (con2 != null && con2.IsConnected)
                {
                    foreach (Connector item in con2.AllRefs)
                    {
                        if (item != null && item.IsConnectedTo(con2))
                        {
                            con2.DisconnectFrom(item);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Lấy connector có cao độ cao hơn
        /// </summary>
        /// <param name="pipe"></param>
        /// <returns></returns>
        public static Connector GetConnectorValid1(MEPCurve pipe)
        {
            if (pipe != null && pipe.Location is LocationCurve locationCurve)
            {
                double slope = Math.Round((double)Common.GetValueParameterByBuilt(pipe, BuiltInParameter.RBS_PIPE_SLOPE), 5);

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
                    Connector retval = (conSt.Origin.Z > conEnd.Origin.Z) ? conSt : conEnd;
                    if (retval.IsConnected)
                        return (conSt.Origin.Z < conEnd.Origin.Z) ? conSt : conEnd;

                    return retval;
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
        public static bool GetLenghtElbowFitting45(Document doc, MEPCurve pipeHorizontal, XYZ vectorMoveZ
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
                Connector conHor1 = GetConnectorValid1(pipeHorizontal);
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
        public static bool ConnectPipeVerticalElbow45(Document doc, MEPCurve pipe1, MEPCurve pipe2, bool isHubMode)
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

                        if (!GetLenghtElbowFitting45(doc, pipeHorizontal, vectoMoveZ, out double lengthElbow1, out double distanceCenterToConnector1))
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

        public static Connector GetConnectorClosestTo1(
      Element e,
      XYZ p)
        {
            ConnectorManager cm = GetConnectorManager(e);

            return null == cm
              ? null
              : GetConnectorClosestTo1(cm.Connectors, p);
        }

        private static Connector GetConnectorClosestTo1(
      ConnectorSet connectors,
      XYZ p)
        {
            Connector targetConnector = null;
            double minDist = double.MaxValue;

            foreach (Connector c in connectors)
            {
                if (c.IsConnected)
                    continue;
                double d = c.Origin.DistanceTo(p);

                if (d < minDist)
                {
                    targetConnector = c;
                    minDist = d;
                }
            }
            return targetConnector;
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

        public static View3D Get3DView(Document doc)
        {
            FilteredElementCollector collector
              = new FilteredElementCollector(doc);

            collector.OfClass(typeof(View3D));

            foreach (View3D v in collector)
            {
                // skip view templates here because they
                // are invisible in project browsers:

                if (v != null && !v.IsTemplate && v.Name == "{3D}")
                {
                    return v;
                }
            }

            return NewView3D(doc);
        }

        public static View3D NewView3D(Document doc)
        {
            FilteredElementCollector collector
              = new FilteredElementCollector(doc);

            collector.OfClass(typeof(View3D));
            if (collector.ToElementIds().Count == 0)
                return null;

            ElementId viewTypeId = collector.ToElementIds().First();

            View3D view = View3D.CreateIsometric(doc, viewTypeId);

            return view;
        }

        public static List<Element> FindElementsByDirection(Document doc, ElementId ognoreId, ElementId elementTypeId, ElementId systemTypeId, XYZ point, XYZ direction, BuiltInCategory built, bool checkMin, Type type)
        {
            var view3d = Get3DView(doc);

            if (view3d == null)
                return null;

            double fsilon = 5000;
            //Change section box: Extend section box to see all object////////////////////////////////////////////////////////////////////////
            var box = view3d.GetSectionBox();

            var max = new XYZ(box.Max.X + fsilon, box.Max.Y + fsilon, box.Max.Z + fsilon);
            var min = new XYZ(box.Min.X - fsilon, box.Min.Y - fsilon, box.Min.Z - fsilon);

            BoundingBoxXYZ newBox = new BoundingBoxXYZ();
            newBox.Min = min;
            newBox.Max = max;
            view3d.SetSectionBox(newBox);
            Global.UIDoc.Document.Regenerate();

            //////////////////////////////////////////////////////////////////////////
            ElementClassFilter filter = new ElementClassFilter(type);
            ReferenceIntersector referenceIntersector = new ReferenceIntersector(filter, FindReferenceTarget.Face, view3d);
            referenceIntersector.FindReferencesInRevitLinks = true;

            List<Element> elements = new List<Element>();
            Dictionary<Element, double> proximities = new Dictionary<Element, double>();

            var mepCurve = Global.UIDoc.Document.GetElement(ognoreId) as MEPCurve;

            IList<ReferenceWithContext> intersectedReferences = referenceIntersector.Find(point, direction);

            double maxFT = 1000 * Common.mmToFT;

            foreach (ReferenceWithContext referenceWithContext in intersectedReferences)
            {
                Autodesk.Revit.DB.Element referenceElement = null;

                Reference reference = referenceWithContext.GetReference();
                if (reference.LinkedElementId != ElementId.InvalidElementId)
                {
                    continue;
                }
                else
                    referenceElement = Global.UIDoc.Document.GetElement(reference);

                if (referenceElement != null && referenceElement as MEPCurve != null)
                {
                    var mep = referenceElement as MEPCurve;

                    if (CheckValid(mep, mepCurve, elementTypeId, systemTypeId) == false)
                        continue;

                    if (referenceElement.Category != null && referenceElement.Category.Id.IntegerValue == (int)built)
                    {
                        var exist = elements.Find(item => item.Id == referenceElement.Id);
                        if (exist == null && ognoreId != referenceElement.Id)
                        {
                            double distance = referenceWithContext.Proximity;

                            if (checkMin)
                            {
                                //Kiêm tra connector
                                List<Connector> connectors = Common.GetConnectionNearest(referenceElement, mepCurve);
                                if (connectors != null && connectors.Count == 2)
                                {
                                    var c1 = connectors[0];
                                    var c2 = connectors[1];
                                    if (c1 != null && c2 != null)
                                    {
                                        distance = Common.ToPoint2D(c1.Origin).DistanceTo(Common.ToPoint2D(c2.Origin));
                                    }
                                }

                                if (distance > maxFT)
                                    continue;
                            }

                            elements.Add(referenceElement);

                            proximities.Add(referenceElement, distance);
                        }
                    }
                }
            }
            elements.Sort(delegate (Element e1, Element e2)
            {
                var d1 = proximities[e1];
                var d2 = proximities[e2];

                return d1.CompareTo(d2);
            }
            );

            //////////////////////////////////////////////////////////////////////////
            view3d.SetSectionBox(box);
            Global.UIDoc.Document.Regenerate();
            //////////////////////////////////////////////////////////////////////////

            return elements;
        }

        private static bool CheckValid(MEPCurve mep, MEPCurve mepCurve, ElementId elementTypeId, ElementId systemTypeId)
        {
            //Check offset////////////////////////////////////////////////////////////////////////
            if (mep as Duct != null || mep as CableTray != null)
            {
                if (mepCurve != null)
                {
                    var para = mepCurve.get_Parameter(BuiltInParameter.RBS_CURVE_VERT_OFFSET_PARAM);
                    if (para != null)
                    {
                        var value = para.AsInteger();
                        if (value == 1) //Bottom
                        {
                            var offset1 = mepCurve.LookupParameter("Bottom Elevation").AsDouble();
                            var offset2 = mep.LookupParameter("Bottom Elevation").AsDouble();

                            if (Math.Abs(offset1 - offset2) > 0.01)
                            {
                                return false;
                            }
                        }
                        else if (value == 2) //Top
                        {
                            var offset1 = mepCurve.LookupParameter("Top Elevation").AsDouble();
                            var offset2 = mep.LookupParameter("Top Elevation").AsDouble();

                            if (Math.Abs(offset1 - offset2) > 0.01)
                            {
                                return false;
                            }
                        }
                        else
                        {
                            var offset1 = GetOffset(mepCurve);// mepCurve.LookupParameter("Offset").AsDouble();
                            var offset2 = GetOffset(mep);// mep.LookupParameter("Offset").AsDouble();

                            if (Math.Abs(offset1 - offset2) > 0.01)
                            {
                                return false;
                            }
                        }
                    }
                }
            }

            //////////////////////////////////////////////////////////////////////////

            //Check element type and system type

            if (mepCurve != null)
            {
                if (mep.GetTypeId() != mepCurve.GetTypeId())
                    return false;
            }

            if (systemTypeId != ElementId.InvalidElementId)
            {
                ElementId systemId = ElementId.InvalidElementId;
                if (mep as Pipe != null)
                {
                    systemId = mep.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsElementId();
                }
                else if (mep as Duct != null)
                {
                    systemId = mep.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM).AsElementId();
                }
                else
                {
                    systemId = mep.get_Parameter(BuiltInParameter.RBS_CTC_SERVICE_TYPE).AsElementId();
                }

                if (systemId != systemTypeId)
                    return false;
            }

            return true;
        }

        public static double GetOffset(Element element)
        {
            return element.LookupParameter("Middle Elevation").AsDouble();
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