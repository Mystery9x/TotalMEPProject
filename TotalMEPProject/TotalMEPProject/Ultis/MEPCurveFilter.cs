using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI.Selection;
using System;

namespace TotalMEPProject.Ultis
{
    public class MEPCurveFilter : ISelectionFilter
    {
        private Type _Type = null;
        public XYZ _Point = null;
        public Element _Element = null;

        public MEPCurveFilter(Type type = null)
        {
            _Type = type;
        }

        public bool AllowElement(Element element)
        {
            if (element == null)
                return false;

            if (element.Category == null)
                return false;

            if (_Type == null)
            {
                if ((BuiltInCategory)element.Category.Id.IntegerValue != BuiltInCategory.OST_PipeCurves &&
                    (BuiltInCategory)element.Category.Id.IntegerValue != BuiltInCategory.OST_DuctCurves &&
                    (BuiltInCategory)element.Category.Id.IntegerValue != BuiltInCategory.OST_CableTray &&
                    (BuiltInCategory)element.Category.Id.IntegerValue != BuiltInCategory.OST_Conduit)
                    return false;
            }
            else
            {
                if (typeof(Pipe) == _Type)
                {
                    if ((BuiltInCategory)element.Category.Id.IntegerValue != BuiltInCategory.OST_PipeCurves)
                        return false;
                }
                else if (typeof(Duct) == _Type)
                {
                    if ((BuiltInCategory)element.Category.Id.IntegerValue != BuiltInCategory.OST_DuctCurves)
                        return false;
                }
                else if (typeof(CableTray) == _Type)
                {
                    if ((BuiltInCategory)element.Category.Id.IntegerValue != BuiltInCategory.OST_CableTray)
                        return false;
                }
                else if (typeof(Conduit) == _Type)
                {
                    if ((BuiltInCategory)element.Category.Id.IntegerValue != BuiltInCategory.OST_Conduit)
                        return false;
                }
            }
            _Element = element;
            return true;
        }

        public bool AllowReference(Reference refer, XYZ point)
        {
            if (_Element != null)
                _Point = point;
            return true;
        }
    }
}