using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;
using System.Collections.Generic;
using System.Linq;

namespace TotalMEPProject.SelectionFilters
{
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
}