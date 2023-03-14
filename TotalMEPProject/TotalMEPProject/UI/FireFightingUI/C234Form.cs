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
    public partial class C234Form : System.Windows.Forms.Form
    {
        private Request.Request m_request;

        private Request.RequestHandler m_handler;

        private ExternalEvent m_exEvent;

        #region Variable & Properties C2

        public bool isConnectTee = false;
        public bool isConnectNipple = false;
        public bool isElbow = false;

        public FamilySymbol fmlNipple = null;

        public double PipeSizeC2
        {
            get
            {
                var value = cboC2PypeSize.SelectedItem.ToString();

                value = value.Replace(" mm", "");

                double d = 0;
                if (double.TryParse(value, out d) == false)
                    return double.MaxValue;

                return d;
            }
        }

        #endregion Variable & Properties C2

        #region Variable & Properties C3

        public bool isTeeTap = false;

        public double PipeSizeC3
        {
            get
            {
                var value = cboC3PipeSize.SelectedItem.ToString();

                value = value.Replace(" mm", "");

                double d = 0;
                if (double.TryParse(value, out d) == false)
                    return double.MaxValue;

                return d;
            }
        }

        public double Height_
        {
            get
            {
                double height = 0;

                if (double.TryParse(txbC3Length.Text.Trim(), out height) == true)
                {
                    return height;
                }
                return double.MinValue;
            }
        }

        #endregion Variable & Properties C3

        #region Variable & Properties C4

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

        public double PipeSizeC4
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

        #endregion Variable & Properties C4

        public C234Form(ExternalEvent exEvent, Request.RequestHandler handler)
        {
            InitializeComponent();

            m_handler = handler;
            m_exEvent = exEvent;
        }

        private void btnApply_Click(object sender, EventArgs e)
        {
            if (tabControl.SelectedTab.Name == "tabC2SprinklerUp")
            {
                isConnectTee = false;
                isConnectNipple = false;
                //fmlNipple = null;
                if (PipeSizeC2 == double.MaxValue)
                    return;

                if (chkC2ConnectTee.Enabled)
                    isConnectTee = chkC2ConnectTee.Checked;

                if (chkC2Nipple.Enabled)
                    isConnectNipple = chkC2Nipple.Checked;
                isElbow = rdnC2Elbow.Checked;
                if (chkC2Nipple.Checked)
                {
                    fmlNipple = cboC2Nipple.SelectedItem as FamilySymbol;
                }

                AppUtils.sa(cboC2PipeType);
                AppUtils.sa(cboC2PypeSize);
                AppUtils.sa(cboC2Nipple);
                SetFocus();
                MakeRequest(RequestId.SprinklerUp_Aplly);
            }
            else if (tabControl.SelectedTab.Name == "tabC3SprinklerDown")
            {
                if (Height_ == double.MinValue)
                    return;
                isTeeTap = rdnC3TeeTap.Checked;
                AppUtils.sa(cboC3PipeType);
                AppUtils.sa(cboC3PipeSize);
                AppUtils.sa(rdnC3Type1);
                AppUtils.sa(rdnC3Type2);
                AppUtils.sa(rdnC3Type3);
                AppUtils.sa(txbC3Length, txbC3Length.Text.ToString());
                if (rdnC3Type1.Checked)
                {
                    SetFocus();
                    MakeRequest(RequestId.SprinklerDownType1_RUN);
                }
                else if (rdnC3Type2.Checked)
                {
                    SetFocus();
                    MakeRequest(RequestId.SprinklerDownType2_RUN);
                }
                else if (rdnC3Type3.Checked)
                {
                    SetFocus();
                    MakeRequest(RequestId.SprinklerDownType3_RUN);
                }
            }
            else if (tabControl.SelectedTab.Name == "tabC4FlexSprinkler")
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
        }

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

        #endregion Method

        #region Event C2

        private void cboC2PipeType_SelectedIndexChanged(object sender, EventArgs e)
        {
            var familyTypeId = (cboC2PipeType.SelectedItem as ObjectItem).ObjectId;

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
                    AddDiameterC2(CurrentSegment);
            }
        }

        private void rdnC2Elbow_CheckedChanged(object sender, EventArgs e)
        {
            if (rdnC2Elbow.Checked)
                chkC2ConnectTee.Enabled = false;
            else
                chkC2ConnectTee.Enabled = true;
        }

        private void chkC2ConnectTee_CheckedChanged(object sender, EventArgs e)
        {
            if (chkC2ConnectTee.Checked)
            {
                chkC2Nipple.Enabled = true;
                cboC2Nipple.Enabled = true;
            }
            else
            {
                chkC2Nipple.Enabled = false;
                cboC2Nipple.Enabled = false;
            }
        }

        #endregion Event C2

        #region Method C2

        private void AddDiameterC2(Segment segment)
        {
            cboC2PypeSize.Items.Clear();

            foreach (MEPSize size in segment.GetSizes())
            {
                var value = Common.FeetToMmString(size.NominalDiameter) + " mm";

                cboC2PypeSize.Items.Add(value);
            }

            AppUtils.ff(cboC2PypeSize);

            if (cboC2PypeSize.SelectedItem == null && cboC2PypeSize.Items.Count != 0)
                cboC2PypeSize.SelectedIndex = 0;
        }

        private void AddFamilyTypeC2()
        {
            cboC2PipeType.Items.Clear();
            FilteredElementCollector pipeTypes = new FilteredElementCollector(Global.UIDoc.Document).OfClass(typeof(PipeType));
            foreach (MEPCurveType type in pipeTypes)
            {
                ObjectItem item = new ObjectItem(type.Name, type.Id);
                cboC2PipeType.Items.Add(item);
            }

            AppUtils.ff(cboC2PipeType, null);

            if (cboC2PipeType.SelectedItem == null && cboC2PipeType.Items.Count != 0)
                cboC2PipeType.SelectedIndex = 0;
        }

        private void AddNippleC2()
        {
            cboC2Nipple.Items.Clear();
            var lstFmlNipple = new FilteredElementCollector(Global.UIDoc.Document).OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .Where(x => x.Category.Id.IntegerValue == (int)BuiltInCategory.OST_PipeFitting)
                .ToList();

            foreach (FamilySymbol fmlNipple in lstFmlNipple)
            {
                cboC2Nipple.Items.Add(fmlNipple);
            }

            cboC2Nipple.DisplayMember = "FamilyName";

            AppUtils.ff(cboC2Nipple, null);
            if (cboC2Nipple.SelectedItem == null && cboC2Nipple.Items.Count != 0)
                cboC2Nipple.SelectedIndex = 0;
        }

        #endregion Method C2

        #region Event C3

        private void txbC3Length_KeyPress(object sender, KeyPressEventArgs e)
        {
            Common.NumberCheck(sender, e, false);
        }

        private void rdnC3Type1_CheckedChanged(object sender, EventArgs e)
        {
            if (rdnC3Type1.Checked)
                txbC3Length.Enabled = true;
            else
                txbC3Length.Enabled = false;
        }

        private void cboC3PipeType_SelectedIndexChanged(object sender, EventArgs e)
        {
            var familyTypeId = (cboC3PipeType.SelectedItem as ObjectItem).ObjectId;

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
                    AddDiameterC3(CurrentSegment);
            }
        }

        #endregion Event C3

        #region Method C3

        private void AddFamilyTypeC3()
        {
            cboC3PipeType.Items.Clear();
            FilteredElementCollector pipeTypes = new FilteredElementCollector(Global.UIDoc.Document).OfClass(typeof(PipeType));
            foreach (MEPCurveType type in pipeTypes)
            {
                ObjectItem item = new ObjectItem(type.Name, type.Id);
                cboC3PipeType.Items.Add(item);
            }

            AppUtils.ff(cboC3PipeType, null);

            if (cboC3PipeType.SelectedItem == null && cboC3PipeType.Items.Count != 0)
                cboC3PipeType.SelectedIndex = 0;
        }

        private void AddDiameterC3(Segment segment)
        {
            cboC3PipeSize.Items.Clear();

            foreach (MEPSize size in segment.GetSizes())
            {
                var value = Common.FeetToMmString(size.NominalDiameter) + " mm";

                cboC3PipeSize.Items.Add(value);
            }

            AppUtils.ff(cboC3PipeSize);

            if (cboC3PipeSize.SelectedItem == null && cboC3PipeSize.Items.Count != 0)
                cboC3PipeSize.SelectedIndex = 0;
        }

        #endregion Method C3

        #region Event C4

        private void rbC4Type1_CheckedChanged(object sender, EventArgs e)
        {
            if (rbC4Type1.Checked)
            {
                tbC4L1.Enabled = false;
                tbC4L.Enabled = true;
                tbC4L2.Enabled = true;
            }
        }

        private void rbC4Type2_CheckedChanged(object sender, EventArgs e)
        {
            if (rbC4Type2.Checked)
            {
                tbC4L1.Enabled = true;
                tbC4L.Enabled = true;
                tbC4L2.Enabled = true;
            }
        }

        private void rbC4Type3_CheckedChanged(object sender, EventArgs e)
        {
            if (rbC4Type3.Checked)
            {
                tbC4L1.Enabled = false;
                tbC4L.Enabled = true;
                tbC4L2.Enabled = false;
            }
        }

        private void cboC4PipeType_SelectedIndexChanged(object sender, EventArgs e)
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
                    AddDiameterC4(CurrentSegment);
            }
        }

        #endregion Event C4

        #region MethodC4

        private void AddFamilyTypeC4()
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

        private void AddDiameterC4(Segment segment)
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

        #endregion MethodC4

        private void C234Form_Load(object sender, EventArgs e)
        {
            //C2
            AddFamilyTypeC2();
            AddNippleC2();

            rdnC2Tee.Checked = true;
            chkC2ConnectTee.Checked = false;
            chkC2Nipple.Enabled = false;
            cboC2Nipple.Enabled = false;

            //C3
            rdnC3TeeTap.Checked = true;

            rdnC3Type1.Checked = true;
            AppUtils.ff(txbC3Length, "200");

            AppUtils.ff(rdnC3Type1);
            AppUtils.ff(rdnC3Type2);
            AppUtils.ff(rdnC3Type3);
            AddFamilyTypeC3();

            //C4
            AddFamilyTypeC4();

            rbC4Type1.Checked = true;
            rdnC4TeeTap.Checked = true;

            AppUtils.ff(rbC4Type1);
            AppUtils.ff(rbC4Type2);
            AppUtils.ff(rbC4Type3);
            AppUtils.ff(tbC4L1);
            AppUtils.ff(tbC4L);
            AppUtils.ff(tbC4L2);
        }

        private void btnOK_Click(object sender, EventArgs e)
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
            this.Close();
        }
    }
}