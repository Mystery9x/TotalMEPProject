using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Windows;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TotalMEPProject.Request;
using TotalMEPProject.Services;
using TotalMEPProject.Ultis;

namespace TotalMEPProject.UI.FireFightingUI
{
    public partial class FlexSprinklerForm : System.Windows.Forms.Form
    {
        #region Variable

        private Request.RequestHandler m_handler;

        private ExternalEvent m_exEvent;

        #endregion Variable

        #region Properties

        public bool IsCheckedTee
        {
            get => rdnC4TeeTap.Checked;
        }

        public bool IsCheckedType1
        {
            get => rbC4Type1.Checked;
        }

        public bool IsCheckedType2
        {
            get => rbC4Type2.Checked;
        }

        public bool IsCheckedType3
        {
            get => rbC4Type3.Checked;
        }

        public ElementId FamilyType
        {
            get
            {
                return (cboC4PipeType.SelectedItem as ObjectItem).ObjectId;
            }
        }

        public double PipeSize
        {
            get
            {
                var value = cboC4PipeSize.SelectedItem.ToString();

                value = value.Replace(" mm", "");

                double d = 0;
                if (double.TryParse(value, out d) == false)
                    return double.MaxValue;

                return d;
            }
        }

        public double VerticalPipeLengthL2
        {
            get
            {
                double dHeight = 0;

                if (double.TryParse(tbC4L2.Text.Trim(), out dHeight) == true)
                {
                    return dHeight;
                }
                return double.MinValue;
            }
        }

        public double HorizontalPipeLengthL
        {
            get
            {
                double dHeight = 0;

                if (double.TryParse(tbC4L.Text.Trim(), out dHeight) == true)
                {
                    return dHeight;
                }
                return double.MinValue;
            }
        }

        public double ExtendPipeLengthL1
        {
            get
            {
                double dHeight = 0;

                if (double.TryParse(tbC4L1.Text.Trim(), out dHeight) == true)
                {
                    return dHeight;
                }
                return double.MinValue;
            }
        }

        #endregion Properties

        #region Constructor

        public FlexSprinklerForm(ExternalEvent exEvent, Request.RequestHandler handler)
        {
            InitializeComponent();
            m_handler = handler;
            m_exEvent = exEvent;
        }

        #endregion Constructor

        #region Method

        public void PressCancel(int count = 2)
        {
            IWin32Window _revit_window = new WindowHandle(ComponentManager.ApplicationWindow);

            for (int i = 0; i < count; i++)
            {
                Press.PostMessage(_revit_window.Handle, (uint)Press.KEYBOARD_MSG.WM_KEYDOWN, (uint)Keys.Escape, 0);
            }
        }

        private void SetFocus()
        {
            IntPtr hBefore = DisplayService.GetForegroundWindow();
            DisplayService.SetForegroundWindow(ComponentManager.ApplicationWindow);
        }

        public void MakeRequest(RequestId request)
        {
            m_handler.Request.Make(request);
            m_exEvent.Raise();
        }

        private void AddFamilyType()
        {
            cboC4PipeType.Items.Clear();
            FilteredElementCollector pipeTypes = new FilteredElementCollector(Global.UIDoc.Document).OfClass(typeof(PipeType));
            foreach (MEPCurveType type in pipeTypes)
            {
                ObjectItem item = new ObjectItem(type.Name, type.Id);
                cboC4PipeType.Items.Add(item);
            }

            AppUtils.ff(cboC4PipeType, null);

            if (cboC4PipeType.SelectedItem == null && cboC4PipeType.Items.Count != 0)
                cboC4PipeType.SelectedIndex = 0;
        }

        private void AddDiameter(Segment segment)
        {
            cboC4PipeSize.Items.Clear();

            foreach (MEPSize size in segment.GetSizes())
            {
                var value = Common.FeetToMmString(size.NominalDiameter) + " mm";

                cboC4PipeSize.Items.Add(value);
            }

            AppUtils.ff(cboC4PipeSize);

            if (cboC4PipeSize.SelectedItem == null && cboC4PipeSize.Items.Count != 0)
                cboC4PipeSize.SelectedIndex = 0;
        }

        #endregion Method

        #region Event

        private void FlexSprinklerForm_Load(object sender, EventArgs e)
        {
            AddFamilyType();

            rbC4Type1.Checked = true;
            rdnC4TeeTap.Checked = true;

            AppUtils.ff(rbC4Type1);
            AppUtils.ff(rbC4Type2);
            AppUtils.ff(rbC4Type3);
            AppUtils.ff(tbC4L1);
            AppUtils.ff(tbC4L);
            AppUtils.ff(tbC4L2);
        }

        private void btnC4Run_Click(object sender, EventArgs e)
        {
            if (VerticalPipeLengthL2 == double.MinValue && tbC4L2.Enabled)
                return;

            if (ExtendPipeLengthL1 == double.MinValue && tbC4L1.Enabled)
                return;

            if (HorizontalPipeLengthL == double.MinValue && tbC4L.Enabled)
                return;
            AppUtils.sa(cboC4PipeType);
            AppUtils.sa(cboC4PipeSize);
            AppUtils.sa(rbC4Type1);
            AppUtils.sa(rbC4Type2);
            AppUtils.sa(rbC4Type3);
            AppUtils.sa(tbC4L1);
            AppUtils.sa(tbC4L);
            AppUtils.sa(tbC4L2);

            MakeRequest(RequestId.FlexSprinker_RUN);
        }

        private void btnC4OK_Click(object sender, EventArgs e)
        {
            PressCancel();
            AppUtils.sa(rbC4Type1);
            AppUtils.sa(rbC4Type2);
            AppUtils.sa(rbC4Type3);
            AppUtils.sa(tbC4L1);
            AppUtils.sa(tbC4L);
            AppUtils.sa(tbC4L2);
            AppUtils.sa(cboC4PipeType);
            AppUtils.sa(cboC4PipeSize);
            App.m_flexSprinklerForm = null;
            this.Close();
        }

        #endregion Event

        private void cboC4PipeType_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            var familyTypeId = (cboC4PipeType.SelectedItem as ObjectItem).ObjectId;

            var familyType = Global.UIDoc.Document.GetElement(familyTypeId) as MEPCurveType;

            if (familyType.RoutingPreferenceManager != null)
            {
                PipeSegment CurrentSegment = null;
                int count = familyType.RoutingPreferenceManager.GetNumberOfRules(RoutingPreferenceRuleGroupType.Segments);

                for (int i = 0; i < count; i++)
                {
                    var rule = familyType.RoutingPreferenceManager.GetRule(RoutingPreferenceRuleGroupType.Segments, i);

                    CurrentSegment = Global.UIDoc.Document.GetElement(rule.MEPPartId) as PipeSegment;
                }

                if (CurrentSegment != null)
                    AddDiameter(CurrentSegment);
            }
        }

        private void rbC4Type1_CheckedChanged_1(object sender, EventArgs e)
        {
            if (rbC4Type1.Checked)
            {
                tbC4L1.Enabled = false;
                tbC4L.Enabled = true;
                tbC4L2.Enabled = true;
            }
        }

        private void rbC4Type2_CheckedChanged_1(object sender, EventArgs e)
        {
            if (rbC4Type2.Checked)
            {
                tbC4L1.Enabled = true;
                tbC4L.Enabled = true;
                tbC4L2.Enabled = true;
            }
        }

        private void rbC4Type3_CheckedChanged_1(object sender, EventArgs e)
        {
            if (rbC4Type3.Checked)
            {
                tbC4L1.Enabled = false;
                tbC4L.Enabled = true;
                tbC4L2.Enabled = false;
            }
        }
    }
}