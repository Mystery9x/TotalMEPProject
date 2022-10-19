using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;

namespace TotalMEPProject.Ultis
{
    public class GeometryUtils
    {
        /// <summary>
        /// Check the point is belong to the Line or not
        /// </summary>
        /// <param name="line"></param>
        /// <param name="point"></param>
        /// <returns></returns>
        public static bool IsPointBelongToLine(XYZ point, Line line)
        {
            if (line == null || point == null)
                return false;

            XYZ pointcheck = ScalePoint(point, 100);
            Line linecheck = Line.CreateBound(ScalePoint(line.GetEndPoint(0), 100), ScalePoint(line.GetEndPoint(1), 100));

            double disPointToLine = Math.Round(linecheck.Distance(pointcheck), 3);

            if (disPointToLine == 0)
                return true;
            else
                return false;
        }

        /// <summary>
        /// Multiply the coordinates of the point by the input factor
        /// </summary>
        /// <param name="point"></param>
        /// <param name="a"></param>
        /// <returns></returns>
        private static XYZ ScalePoint(XYZ point, double a)
        {
            return new XYZ(point.X * a, point.Y * a, point.Z * a);
        }

        public static bool IsParallel(XYZ first, XYZ second)
        {
            XYZ product = first.CrossProduct(second);
            double length = product.GetLength();
            return Common.IsEqual(length, 0, 10e-3);
        }

        public static bool IsPerpendicular(XYZ first, XYZ second)
        {
            if (first != null && second != null)
            {
                double product = first.DotProduct(second);
                return Common.IsEqual(product, 0);
            }
            return false;
        }

        /// <summary>
        /// get all solids of a given element
        /// </summary>
        public static List<Solid> GetAllSolids(Document doc,
                                               Element elem,
                                               bool getInsGeo,
                                               View view = null,
                                               bool computeReferences = true,
                                               bool includeNonVisibleObjects = false)
        {
            Options options = new Options
            {
                ComputeReferences = computeReferences,
                IncludeNonVisibleObjects = includeNonVisibleObjects
            };
            if (view != null)
                options.View = view;

            GeometryElement geoElem = elem.get_Geometry(options);
            List<Solid> solids = new List<Solid>();
            GetSolidFromGeometry(doc, geoElem, getInsGeo, ref solids);
            return solids;
        }

        /// <summary>
        /// recursively get solid from geometry element
        /// </summary>
        public static void GetSolidFromGeometry(Document doc, GeometryElement geoElem, bool getInstGeo, ref List<Solid> solids, Autodesk.Revit.DB.View view = null)
        {
            foreach (GeometryObject geoObj in geoElem)
            {
                if (geoObj is Solid solid
                    && solid.Volume > 0
                    && IsSolidGraphicallyVisible(doc, view, solid))
                    solids.Add(solid);
                else if (geoObj is GeometryInstance geoInst)
                {
                    GeometryElement innerGeo = getInstGeo ? geoInst.GetInstanceGeometry() : geoInst.GetSymbolGeometry();
                    GetSolidFromGeometry(doc, innerGeo, getInstGeo, ref solids, view);
                }
            }
        }

        /// <summary>
        /// Determine if soid is visivle in view
        /// </summary>
        public static bool IsSolidGraphicallyVisible(Document doc, Autodesk.Revit.DB.View view, Solid solid)
        {
            if (doc != null
                && view != null
                && solid.GraphicsStyleId != null
                && solid.GraphicsStyleId != ElementId.InvalidElementId)
            {
                if (doc.GetElement(solid.GraphicsStyleId) is GraphicsStyle graphicalStyle
                    && graphicalStyle.GraphicsStyleCategory != null)
                    return graphicalStyle.GraphicsStyleCategory.get_Visible(view);
            }
            return true;
        }

        /// <summary>
        /// filter out all points that are nots vertices of polygon
        /// </summary>
        /// <param name="vertices"></param>
        /// <returns></returns>
        public static List<XYZ> MergePoints(List<XYZ> vertices)
        {
            List<XYZ> mergedPoints = new List<XYZ>();

            if (vertices?.Count > 0)
            {
                vertices = Distinct(vertices);
                List<XYZ> pointsOnTheSameEdge = new List<XYZ>();
                int baseIndex = GetStartIndex(vertices);

                for (int currentIndex = 0; currentIndex <= vertices.Count; currentIndex++)
                {
                    XYZ prePoint = pointsOnTheSameEdge.LastOrDefault();
                    XYZ currentPoint = vertices[baseIndex];
                    baseIndex = (baseIndex + 1) % vertices.Count;
                    XYZ nextPoint = vertices[baseIndex];

                    pointsOnTheSameEdge.Add(currentPoint);
                    if (pointsOnTheSameEdge.Count > 1)
                    {
                        XYZ currentDir = (currentPoint - prePoint).Normalize();
                        XYZ nextDir = (nextPoint - currentPoint).Normalize();

                        if (!Common.IsEqual(currentDir, nextDir))
                        {
                            mergedPoints.Add(pointsOnTheSameEdge.First());
                            mergedPoints.Add(pointsOnTheSameEdge.Last());
                            pointsOnTheSameEdge.Clear();
                            pointsOnTheSameEdge.Add(currentPoint);
                        }
                    }
                }
                mergedPoints = Distinct(mergedPoints);
            }
            return mergedPoints;
        }

        /// <summary>
        /// filter out all duplicated points
        /// </summary>
        /// <param name="points"></param>
        /// <returns></returns>
        public static List<XYZ> Distinct(List<XYZ> points)
        {
            List<XYZ> diffPoints = new List<XYZ>();
            if (points?.Count > 0)
            {
                foreach (XYZ point in points)
                {
                    if (diffPoints.All(x => !Common.IsEqualPoint2D(x, point)))
                        diffPoints.Add(point);
                }
            }
            return diffPoints;
        }

        /// <summary>
        /// get the index of the first point whose
        /// location is at the corner or a polygon
        /// formed by given points
        /// </summary>
        /// <param name="points"></param>
        /// <returns></returns>
        public static int GetStartIndex(List<XYZ> points)
        {
            if (points?.Count > 0)
            {
                int count = points.Count;
                for (int index = 0; index < count; index++)
                {
                    XYZ prePoint = points[(index + count - 1) % count];
                    XYZ currentPoint = points[index];
                    XYZ nextPoint = points[(index + 1) % count];

                    XYZ nextDir = (nextPoint - currentPoint).Normalize();
                    XYZ preDir = (currentPoint - prePoint).Normalize();

                    if (!Common.IsEqual(nextDir, preDir))
                        return index;
                }
            }
            return 0;
        }
    }
}