using Autodesk.Revit.DB;
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