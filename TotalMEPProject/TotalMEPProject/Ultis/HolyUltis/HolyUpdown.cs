using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB.Mechanical;
using System.Windows.Forms;
using System.Xml;
using System.IO;

namespace TotalMEPProject.Ultis.HolyUltis
{
    public class Data
    {
        public MEPCurve _MEPCurve = null;
        public XYZ _Start = null;
        public XYZ _End = null;

        public XYZ _PStart = null;
        public XYZ _PEnd = null;

        public XYZ _UnionPoint01 = null;
        public XYZ _UnionPoint02 = null;

        public Element _Elbow01 = null;
        public Element _Elbow02 = null;
        public Element _Elbow03 = null;
        public Element _Elbow04 = null;

        public MEPCurve _Vertical01 = null;
        public MEPCurve _Vertical02 = null;

        public MEPCurve _Other01 = null;
        public MEPCurve _Other02 = null;

        public double _OldOffset = 0;
        public double _OldOffsetApply = 0;

        public Data(MEPCurve mepCurve)
        {
            _MEPCurve = mepCurve;
        }
    }
}