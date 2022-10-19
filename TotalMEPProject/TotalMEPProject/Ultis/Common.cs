using Autodesk.Revit.DB;
using System;
using System.Linq;

namespace TotalMEPProject.Ultis
{
    public class Common
    {
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
}