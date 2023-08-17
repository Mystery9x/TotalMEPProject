using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Visual;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using TotalMEPProject.Ultis;

namespace TotalMEPProject.SelectionFilters
{
    public class PipeSelectionFilter : ISelectionFilter
    {
        private List<ElementId> _exceptIds = new List<ElementId>();

        public PipeSelectionFilter(List<ElementId> exceptIds = null)
        {
            if (exceptIds != null)
                _exceptIds = exceptIds;
        }

        public bool AllowElement(Element elem)
        {
            if (elem is Pipe /*|| (elem is FamilyInstance instance && instance.MEPModel != null)*/)
            {
                if (!_exceptIds.Contains(elem.Id))
                    return true;
                else
                    return false;
            }
            else
                return false;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return true;
        }
    }

    public class ElementSelectionFilters : ISelectionFilter
    {
        private Document _doc;

        public ElementSelectionFilters(Document doc)
        {
            _doc = doc;
        }

        public bool AllowElement(Element element)
        {
            return true;
        }

        public bool AllowReference(Reference reference, XYZ point)
        {
            var elem = _doc.GetElement(reference);
            if (elem == null || elem.Category == null)
                return false;

            //if (elem.Category.Id.IntegerValue == (int)BuiltInCategory.OST_PipeCurves ||
            //    elem.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Conduit ||
            //    elem.Category.Id.IntegerValue == (int)BuiltInCategory.OST_DuctCurves ||
            //    elem.Category.Id.IntegerValue == (int)BuiltInCategory.OST_CableTray)
            //{
            //    var locationCurve = elem.Location as LocationCurve;
            //    var line = locationCurve.Curve as Line;
            //    if (GeometryUtils.IsPointBelongToLine(point, line))
            //        return true;
            //}
            return false;
        }
    }

    public class AdditionalElementSelectionFilters : ISelectionFilter
    {
        private Document _doc;
        private XYZ _pointCheck;

        public AdditionalElementSelectionFilters(Document doc, XYZ pointCheck)
        {
            _doc = doc;
            _pointCheck = pointCheck;
        }

        public bool AllowElement(Element element)
        {
            return true;
        }

        public bool AllowReference(Reference reference, XYZ point)
        {
            var elem = _doc.GetElement(reference);
            if (elem == null || elem.Category == null)
                return false;

            //if (elem.Category.Id.IntegerValue == (int)BuiltInCategory.OST_PipeCurves ||
            //    elem.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Conduit ||
            //    elem.Category.Id.IntegerValue == (int)BuiltInCategory.OST_DuctCurves ||
            //    elem.Category.Id.IntegerValue == (int)BuiltInCategory.OST_CableTray)
            //{
            //    var locationCurve = elem.Location as LocationCurve;
            //    var line = locationCurve.Curve as Line;

            //    if (GeometryUtils.IsPointBelongToLine(_pointCheck, line) &&
            //        GeometryUtils.IsPointBelongToLine(point, line))
            //        return true;
            //}
            return false;
        }
    }

    public class HostSelectionFilters : ISelectionFilter
    {
        private List<ElementId> _hostCategoryIds;

        public HostSelectionFilters(List<ElementId> hostCategoryIds)
        {
            _hostCategoryIds = hostCategoryIds;
        }

        public bool AllowElement(Element element)
        {
            if (element.Category != null &&
                _hostCategoryIds.Any(x => x.IntegerValue == element.Category.Id.IntegerValue))
                return true;

            return false;
        }

        public bool AllowReference(Reference reference, XYZ point)
        {
            return true;
        }
    }

    public class HostLinkSelectionFilters : ISelectionFilter
    {
        private Document _doc;
        private List<ElementId> _hostCategoryIds;

        public HostLinkSelectionFilters(Document doc, List<ElementId> hostCategoryIds)
        {
            _doc = doc;
            _hostCategoryIds = hostCategoryIds;
        }

        public bool AllowElement(Element element)
        {
            return true;
        }

        public bool AllowReference(Reference reference, XYZ point)
        {
            Element elem = _doc.GetElement(reference);
            if (elem != null &&
                elem is RevitLinkInstance linkInstance)
            {
                Document linkDoc = linkInstance.GetLinkDocument();
                Element linkElement = linkDoc.GetElement(reference.LinkedElementId);
                if (linkElement != null &&
                    linkElement.Category != null &&
                    _hostCategoryIds.Any(x => x.IntegerValue == linkElement.Category.Id.IntegerValue))
                    return true;
            }

            return false;
        }
    }

    public class DetailCurveItem
    {
        public Autodesk.Revit.DB.PolyLine Polyline = null;

        public XYZ P0 = null;
        public XYZ P1 = null;

        public DetailCurve DetailCurve = null;

        public DetailCurveItem(Autodesk.Revit.DB.PolyLine pline, XYZ p0, XYZ p1, DetailCurve curve)
        {
            Polyline = pline;
            P0 = p0;
            P1 = p1;
            DetailCurve = curve;
        }
    }

    public class PickLineFiler : ISelectionFilter
    {
        //Dictionary contains Reference with Cursor position
        public Dictionary<ElementId, XYZ> PickElements = new Dictionary<ElementId, XYZ>();

        public DetailCurve TempCurve = null;

        public Curve Curve_selected = null;

        public XYZ Start = null;
        public XYZ End = null;

        public LinePatternElement LinePatternElement = null;

        private Color Backup_color = null;
        private ElementId Backup_Pattern = ElementId.InvalidElementId;

        private Category LineStyle = null;

        private bool setNewStyle = false;

        private Element LastElement = null;

        public List<DetailCurveItem> DetailCurveItems = new List<DetailCurveItem>();

        public PickLineFiler()
        {
        }

        public bool AllowElement(Element elem)
        {
            if (elem != null && TempCurve != null && elem.Id == TempCurve.Id)
                return false;

            if (LastElement != elem || elem.Category == null)
                RemoveTemp();

            if (elem.Category == null)
                return false;

            LastElement = elem;

            return true;
        }

        public void RemoveDetailCurveItems()
        {
            if (DetailCurveItems.Count != 0)
            {
                foreach (var detailCurve in DetailCurveItems)
                {
                    var detail = detailCurve.DetailCurve;

                    if (detail != null)
                    {
                        try
                        {
                            Global.UIDoc.Document.Delete(detail.Id);
                        }
                        catch (System.Exception ex)
                        {
                        }
                    }
                }
            }
        }

        public void RemoveTemp()
        {
            if (TempCurve != null)
            {
                try
                {
                    if (setNewStyle == true)
                    {
                        //Reset backup value
                        if (Backup_color != null && Backup_Pattern != ElementId.InvalidElementId)
                        {
                            GraphicsStyle gs = TempCurve.LineStyle as GraphicsStyle;

                            if (LinePatternElement != null && gs != null)
                            {
                                gs.GraphicsStyleCategory.LineColor = Backup_color;
                                gs.GraphicsStyleCategory.SetLinePatternId(Backup_Pattern, GraphicsStyleType.Projection);
                            }
                        }
                    }

                    Global.UIDoc.Document.Delete(TempCurve.Id);
                    TempCurve = null;
                }
                catch (System.Exception ex)
                {
                }
            }
        }

        public bool AllowReference(Reference refer, XYZ position)
        {
            bool result = false;
            try
            {
                var element = LastElement;

                if (element == null)
                    return false;

                var e = element.GetGeometryObjectFromReference(refer);

                if (e != null)
                {
                    if (e is Curve)
                    {
                        Curve_selected = e as Curve;
                        result = true;
                    }
                    else if (e is PlanarFace)
                    {
                        //result =  true;
                    }
                    else if (e is Edge)
                    {
                        Curve_selected = (e as Edge).AsCurve();
                        result = true;
                    }
                    else if (e is Line)
                    {
                        Curve_selected = e as Line;
                        result = true;
                    }
                    else if (e is Autodesk.Revit.DB.PolyLine)
                    {
                        var poly = e as Autodesk.Revit.DB.PolyLine;

                        for (int i = 0; i < poly.NumberOfCoordinates; i++)
                        {
                            var next = (i + 1) % poly.NumberOfCoordinates;
                            if (next == 0)
                                continue;

                            var p0 = poly.GetCoordinate(i);

                            var p1 = poly.GetCoordinate(next);

                            if (p0.DistanceTo(p1) < Common._sixteenth)
                                continue;

                            try
                            {
                                var line2 = Line.CreateBound(p0, p1);
                                if (CheckPoint(line2, position) == true)
                                {
                                    var find = DetailCurveItems.Find(item => item.Polyline == poly && Common.Compare(item.P0, p0) == 0 && Common.Compare(item.P1, p1) == 0);

                                    if (find == null)
                                    {
                                        DetailCurve detailCurve = Global.UIDoc.Document.Create.NewDetailCurve(Global.UIDoc.Document.ActiveView, line2);

                                        DetailCurveItem item = new DetailCurveItem(poly, p0, p1, detailCurve);
                                        DetailCurveItems.Add(item);

                                        Curve_selected = detailCurve.GetCurve();

                                        //TempCurve = detailCurve;
                                    }
                                    else
                                    {
                                        Curve_selected = find.DetailCurve.GetCurve();

                                        //TempCurve = find.DetailCurve;
                                    }

                                    //var modelCurve = Common.CreateModelLine(p0, p1);

                                    //ICollection<ElementId> elementIds = Global.m_uiDoc.Selection.GetElementIds();
                                    //elementIds.Clear();

                                    //elementIds.Add(detailCurve.Id);

                                    //Global.m_uiDoc.Selection.SetElementIds(elementIds);

                                    //return false;

                                    //////////////////////////////////////////////////////////////////////////
                                    result = true;
                                    break;
                                    //////////////////////////////////////////////////////////////////////////
                                }
                            }
                            catch (System.Exception ex)
                            {
                            }
                        }
                        //return true;
                    }
                    else
                    {
                    }
                }
                else
                {
                    if (element is DetailCurve)
                    {
                        Curve_selected = (element as DetailCurve).GetCurve();

                        result = true;
                    }
                    else if (element is ModelLine)
                    {
                        Curve_selected = (element as ModelLine).GetCurve();
                        result = true;
                    }
                }
            }
            catch (System.Exception ex)
            {
            }

            //if (result == true && App.m_PickLineForm != null)
            //{
            //    var windowPoint = GetCursorPoint();

            //    //Common.CreateModelLine(q, new XYZ(q.X + 20, q.Y, q.Z));

            //    if (PickElements.ContainsKey(refer.ElementId) == true)
            //        PickElements.Remove(refer.ElementId);

            //    PickElements.Add(refer.ElementId, windowPoint);

            //    //////////////////////////////////////////////////////////////////////////
            //    Start = Curve_selected.GetEndPoint(0);
            //    End = Curve_selected.GetEndPoint(1);
            //    //////////////////////////////////////////////////////////////////////////

            //    if (App.m_PickLineForm.HorizontalEnum_ == HorizontalEnum.Side)
            //    {
            //        var start = Curve_selected.GetEndPoint(0);
            //        var end = Curve_selected.GetEndPoint(1);

            //        var elevation = Global.UIDoc.ActiveView.GenLevel.Elevation;

            //        start = new XYZ(start.X, start.Y, elevation);
            //        end = new XYZ(end.X, end.Y, elevation);

            //        var Curve_selected_ = Line.CreateBound(start, end);

            //        double doffset = 0;
            //        if (App.m_CreateMEPForm.MEPType_ == MEPType.Pipe)
            //        {
            //            doffset = (App.m_CreateMEPForm.MEPSize_ as MEPSize).OuterDiameter;
            //        }
            //        else if (App.m_CreateMEPForm.MEPType_ == MEPType.Conduit)
            //        {
            //            doffset = (App.m_CreateMEPForm.MEPSize_ as ConduitSize).OuterDiameter;
            //        }
            //        else if (App.m_CreateMEPForm.MEPType_ == MEPType.Round_Duct)
            //        {
            //            doffset = (App.m_CreateMEPForm.MEPSize_ as MEPSize).NominalDiameter/*OuterDiameter*/;
            //        }
            //        else if (App.m_CreateMEPForm.MEPType_ == MEPType.Oval_Duct || App.m_CreateMEPForm.MEPType_ == MEPType.Rectangular_Duct)
            //        {
            //            doffset = App.m_CreateMEPForm.MEP_Width * Common.mmToFT;
            //        }
            //        else
            //        {
            //            doffset = App.m_CreateMEPForm.MEP_Width * Common.mmToFT;
            //        }
            //        var offsets = Common.CreateOffsetCurve(Curve_selected_, doffset, true);

            //        //Get line with fatherest point
            //        Curve curve1 = offsets[0];
            //        Curve curve2 = offsets[1];

            //        var d1 = curve1.Distance(windowPoint);
            //        var d2 = curve2.Distance(windowPoint);

            //        if (TempCurve == null)
            //        {
            //            try
            //            {
            //                if (d1 </*>*/ d2)
            //                {
            //                    TempCurve = Global.UIDoc.Document.Create.NewDetailCurve(Global.UIDoc.Document.ActiveView, curve1);
            //                }
            //                else
            //                {
            //                    TempCurve = Global.UIDoc.Document.Create.NewDetailCurve(Global.UIDoc.Document.ActiveView, curve2);
            //                }
            //            }
            //            catch (System.Exception ex)
            //            {
            //            }

            //            if (LinePatternElement == null)
            //                LinePatternElement = Common.CreateLineParttern(Global.UIDoc.Document, "Custom_Dash_Space");

            //            //////////////////////////////////////////////////////////////////////////
            //            //App.ShowFormTooltip(Global.m_UiApp);
            //            //////////////////////////////////////////////////////////////////////////

            //            //                         //C1
            //            //                         if(LineStyle == null)
            //            //                             LineStyle = Common.CreateLineStyle(Global.m_uiDoc.Document);
            //            //
            //            //                         LineStyle.SetLinePatternId(LinePatternElement.Id, GraphicsStyleType.Projection);
            //            //
            //            //                         TempCurve.LineStyle = LineStyle.GetGraphicsStyle(GraphicsStyleType.Projection);

            //            //C2

            //            if (setNewStyle == true)
            //            {
            //                GraphicsStyle gs = TempCurve.LineStyle as GraphicsStyle;

            //                if (LinePatternElement != null && gs != null)
            //                {
            //                    //Backup
            //                    Backup_color = gs.GraphicsStyleCategory.LineColor;
            //                    Backup_Pattern = gs.GraphicsStyleCategory.GetLinePatternId(GraphicsStyleType.Projection);

            //                    gs.GraphicsStyleCategory.LineColor = new Color(0, 128, 255/*255, 0, 0*/);
            //                    gs.GraphicsStyleCategory.SetLinePatternId(LinePatternElement.Id, GraphicsStyleType.Projection);
            //                }
            //            }

            //            //                         //C3
            //            //                          //Color
            //            //                          OverrideGraphicSettings ogs = new OverrideGraphicSettings();
            //            //                          ogs.SetProjectionLineColor(new Color(0, 255, 0));
            //            //                         if (LinePatternElement != null )
            //            //                             ogs.SetProjectionLinePatternId(LinePatternElement.Id);
            //            //                          ogs.SetProjectionLineWeight(8);
            //            //                          Global.m_uiDoc.ActiveView.SetElementOverrides(TempCurve.Id, ogs);

            //            Global.UIDoc.RefreshActiveView();
            //        }
            //    }
            //}

            return result;
        }

        private bool CheckPoint(Line line2, XYZ position)
        {
            try
            {
                double fsilon = 10;
                var p0 = new XYZ(position.X, position.Y, position.Z - fsilon);
                var p1 = new XYZ(position.X, position.Y, position.Z + fsilon);

                var line = Line.CreateBound(p0, p1);

                IntersectionResultArray resultArray = new IntersectionResultArray();

                var result2 = line2.Intersect(line, out resultArray);

                if (result2 == SetComparisonResult.Overlap)
                    return true;
                else
                {
                    //https://stackoverflow.com/questions/7050186/find-if-point-lays-on-line-segment
                    var x1 = line2.GetEndPoint(0).X;
                    var x2 = line2.GetEndPoint(1).X;
                    var x = position.X;

                    var y1 = line2.GetEndPoint(0).Y;
                    var y2 = line2.GetEndPoint(1).Y;
                    var y = position.Y;

                    var z1 = 0;// line2.GetEndPoint(0).Z;
                    var z2 = 0;// line2.GetEndPoint(1).Z;
                    var z = 0;// position.Z;

                    var AB = Math.Sqrt((x2 - x1) * (x2 - x1) + (y2 - y1) * (y2 - y1) + (z2 - z1) * (z2 - z1));
                    var AP = Math.Sqrt((x - x1) * (x - x1) + (y - y1) * (y - y1) + (z - z1) * (z - z1));
                    var PB = Math.Sqrt((x2 - x) * (x2 - x) + (y2 - y) * (y2 - y) + (z2 - z) * (z2 - z));
                    if (AB == AP + PB)
                        return true;

                    //////////////////////////////////////////////////////////////////////////
                    //var project = line2.Project(position);
                    //if (project == null)
                    //    return false;

                    //var d = project.Distance;
                    //var pp = project.XYZPoint;

                    //var d1 = pp.DistanceTo(line2.GetEndPoint(0));
                    //var d2 = pp.DistanceTo(line2.GetEndPoint(1));

                    //if (d1 != 0 && d2 != 0)
                    //    return true;
                    //////////////////////////////////////////////////////////////////////////
                }

                return false;
            }
            catch (System.Exception ex)
            {
                return false;
            }
        }

        private XYZ GetCursorPoint()
        {
            IList<UIView> uiviews = Global.UIDoc.GetOpenUIViews();
            UIView uiview = null;
            foreach (UIView uv in uiviews)
            {
                if (uv.ViewId.Equals(Global.UIDoc.ActiveView.Id))
                {
                    uiview = uv;
                    break;
                }
            }

            Rectangle rect = uiview.GetWindowRectangle();
            IList<XYZ> corners = uiview.GetZoomCorners();

            var p = System.Windows.Forms.Cursor.Position;

            double dx = (double)(p.X - rect.Left)
              / (rect.Right - rect.Left);

            double dy = (double)(p.Y - rect.Bottom)
              / (rect.Top - rect.Bottom);

            XYZ a = corners[0];
            XYZ b = corners[1];
            XYZ v = b - a;

            XYZ q = a
              + dx * v.X * XYZ.BasisX
              + dy * v.Y * XYZ.BasisY;

            return q;
        }
    }

    public enum HorizontalEnum
    {
        Center = 0,
        Side
    }

    public enum VerticalEnum
    {
        Middle = 0,
        Bottom,
        Top
    }

    public enum MEPType
    {
        Pipe = 0,
        Oval_Duct = 1,
        Rectangular_Duct = 2,
        Round_Duct = 3,
        CableTray = 4,
        Conduit = 5,
        Duct = 6,
    }

    public enum Slope
    {
        Down = 0,
        Up,
        Off
    }
}