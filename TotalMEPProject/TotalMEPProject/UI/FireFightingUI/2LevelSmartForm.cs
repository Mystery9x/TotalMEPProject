using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Windows;
using System;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using TotalMEPProject.Request;
using TotalMEPProject.Services;
using TotalMEPProject.Ultis;

namespace TotalMEPProject.UI.FireFightingUI
{
    public partial class _2LevelSmartForm : System.Windows.Forms.Form
    {
        #region Variable & Properties

        private Request.Request m_request;

        private Request.RequestHandler m_handler;

        private ExternalEvent m_exEvent;

        public ElementId FamilyType
        {
            get
            {
                return (cboC1FamilyType.SelectedItem as ObjectItem).ObjectId;
            }
        }

        public FamilySymbol SelectedNippleFamily
        {
            get
            {
                return (cboC1NippleFamily.SelectedItem as FamilySymbol);
            }
        }

        public bool OptionMainBranchElevation_SameElevation
        {
            get
            {
                return rBC1SameElevation.Checked;
            }
        }

        public bool OptionMainBranchElevation_ElevationDifference
        {
            get
            {
                return rBC1ElevationDifference.Checked;
            }
        }

        public bool OptionMainBranchConnection_TeeOrTap
        {
            get
            {
                return rBC1TeeOrTap.Checked;
            }
        }

        public bool OptionMainBranchConnection_Elbow
        {
            get
            {
                return rBC1ElbowConnect.Checked;
            }
        }

        public bool OptionAddNipple
        {
            get
            {
                return ckbNippleCreating.Checked;
            }
        }

        public bool OptionAddElbowConnection
        {
            get
            {
                return ckbC1ElbowConnection.Checked;
            }
        }

        public double PipeSize
        {
            get
            {
                var value = cboC1PipeSize.SelectedItem.ToString();

                value = value.Replace(" mm", "");

                double d = 0;
                if (double.TryParse(value, out d) == false)
                    return double.MaxValue;

                return d;
            }
        }

        #endregion Variable & Properties

        #region Constructor

        public _2LevelSmartForm(ExternalEvent exEvent, Request.RequestHandler handler)
        {
            InitializeComponent();

            m_handler = handler;
            m_exEvent = exEvent;
            rBC1ElevationDifference.Checked = true;
            rBC1TeeOrTap.Checked = true;
            ckbC1ElbowConnection.Checked = false;
            ckbNippleCreating.Checked = false;
            cboC1NippleFamily.Enabled = false;
        }

        private void _2LevelSmartForm_Load(object sender, EventArgs e)
        {
            AddFamilyType();
            AddNipple();
            AppUtils.ff(rBC1ElevationDifference);
            AppUtils.ff(rBC1SameElevation);
            AppUtils.ff(rBC1TeeOrTap);
            AppUtils.ff(rBC1ElbowConnect);
        }

        #endregion Constructor

        #region Method

        private void AddDiameter(Segment segment)
        {
            cboC1PipeSize.Items.Clear();

            foreach (MEPSize size in segment.GetSizes())
            {
                var value = Common.FeetToMmString(size.NominalDiameter) + " mm";

                cboC1PipeSize.Items.Add(value);
            }

            AppUtils.ff(cboC1PipeSize);

            if (cboC1PipeSize.SelectedItem == null && cboC1PipeSize.Items.Count != 0)
                cboC1PipeSize.SelectedIndex = 0;
        }

        private void AddFamilyType()
        {
            cboC1FamilyType.Items.Clear();
            FilteredElementCollector pipeTypes = new FilteredElementCollector(Global.UIDoc.Document).OfClass(typeof(PipeType));
            foreach (MEPCurveType type in pipeTypes)
            {
                ObjectItem item = new ObjectItem(type.Name, type.Id);
                cboC1FamilyType.Items.Add(item);
            }

            AppUtils.ff(cboC1FamilyType, null);

            if (cboC1FamilyType.SelectedItem == null && cboC1FamilyType.Items.Count != 0)
                cboC1FamilyType.SelectedIndex = 0;
        }

        private void AddNipple()
        {
            cboC1NippleFamily.Items.Clear();
            var lstFmlNipple = new FilteredElementCollector(Global.UIDoc.Document).OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .Where(x => x.Category.Id.IntegerValue == (int)BuiltInCategory.OST_PipeFitting)
                .ToList();

            foreach (FamilySymbol fmlNipple in lstFmlNipple)
            {
                cboC1NippleFamily.Items.Add(fmlNipple);
            }

            cboC1NippleFamily.DisplayMember = "FamilyName";

            AppUtils.ff(cboC1NippleFamily, null);
            if (cboC1NippleFamily.SelectedItem == null && cboC1NippleFamily.Items.Count != 0)
                cboC1NippleFamily.SelectedIndex = 0;
        }

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

        #endregion Method

        #region Event

        private void cboC1FamilyType_SelectedIndexChanged(object sender, EventArgs e)
        {
            var familyTypeId = (cboC1FamilyType.SelectedItem as ObjectItem).ObjectId;

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

        private void btnCancel_Click(object sender, EventArgs e)
        {
            PressCancel();
            App._2LevelSmartForm = null;
            AppUtils.sa(cboC1FamilyType);
            AppUtils.sa(cboC1PipeSize);
            AppUtils.sa(cboC1NippleFamily);
            AppUtils.sa(rBC1ElevationDifference);
            AppUtils.sa(rBC1SameElevation);
            AppUtils.sa(rBC1TeeOrTap);
            AppUtils.sa(rBC1ElbowConnect);
            this.Close();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (PipeSize == double.MaxValue)
                return;

            AppUtils.sa(cboC1FamilyType);
            AppUtils.sa(cboC1PipeSize);
            AppUtils.sa(cboC1NippleFamily);
            AppUtils.sa(rBC1ElevationDifference);
            AppUtils.sa(rBC1SameElevation);
            AppUtils.sa(rBC1TeeOrTap);
            AppUtils.sa(rBC1ElbowConnect);
            SetFocus();
            MakeRequest(RequestId.TwoLevelSmart_OK);
        }

        private void ckbNippleCreating_CheckedChanged(object sender, EventArgs e)
        {
            if (ckbNippleCreating.Checked)
            {
                ckbC1ElbowConnection.Enabled = false;
                cboC1NippleFamily.Enabled = true;
            }
            else
            {
                ckbC1ElbowConnection.Enabled = true;
                cboC1NippleFamily.Enabled = false;
            }
        }

        private void ckbC1ElbowConnection_CheckedChanged(object sender, EventArgs e)
        {
            ReDisplay();
        }

        #endregion Event

        private void rBC1SameElevation_CheckedChanged(object sender, EventArgs e)
        {
            ReDisplay();
        }

        private void rBC1ElbowConnect_CheckedChanged(object sender, EventArgs e)
        {
            ReDisplay();
        }

        private void rBC1ElevationDifference_CheckedChanged(object sender, EventArgs e)
        {
            ReDisplay();
        }

        private void rBC1TeeOrTap_CheckedChanged(object sender, EventArgs e)
        {
            ReDisplay();
        }

        private void ReDisplay()
        {
            rBC1ElevationDifference.Enabled = true;
            rBC1SameElevation.Enabled = true;
            rBC1TeeOrTap.Enabled = true;
            rBC1ElbowConnect.Enabled = true;
            cboC1FamilyType.Enabled = true;
            cboC1PipeSize.Enabled = true;
            ckbC1ElbowConnection.Enabled = true;
            ckbNippleCreating.Enabled = true;
            cboC1NippleFamily.Enabled = true;

            if (rBC1SameElevation.Checked == true)
            {
                rBC1TeeOrTap.Enabled = false;
                rBC1ElbowConnect.Enabled = false;
                cboC1FamilyType.Enabled = false;
                cboC1PipeSize.Enabled = false;
                ckbC1ElbowConnection.Enabled = false;
                ckbNippleCreating.Enabled = false;
                cboC1NippleFamily.Enabled = false;
            }

            if (rBC1ElbowConnect.Checked == true)
            {
                cboC1PipeSize.Enabled = false;
                cboC1FamilyType.Enabled = false;
                ckbNippleCreating.Enabled = false;
                cboC1NippleFamily.Enabled = false;
            }

            if (ckbC1ElbowConnection.Checked == true)
            {
                cboC1PipeSize.Enabled = false;
                cboC1FamilyType.Enabled = false;
                ckbNippleCreating.Enabled = false;
                cboC1NippleFamily.Enabled = false;
            }
        }
    }
}