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
    public partial class SprinkerUpForm : System.Windows.Forms.Form
    {
        #region Variable & Properties

        private Request.Request m_request;

        private Request.RequestHandler m_handler;

        private ExternalEvent m_exEvent;

        public bool isD15 = false;
        public bool isConnectTee = false;
        public bool isConnectNipple = false;
        public bool isElbow = false;

        public FamilySymbol fmlNipple = null;

        public ElementId FamilyType
        {
            get
            {
                return (cboC2PipeType.SelectedItem as ObjectItem).ObjectId;
            }
        }

        public double PipeSize
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

        #endregion Variable & Properties

        public SprinkerUpForm(ExternalEvent exEvent, Request.RequestHandler handler)
        {
            InitializeComponent();

            m_handler = handler;
            m_exEvent = exEvent;
        }

        #region Method

        private void AddDiameter(Segment segment)
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

        private void AddFamilyType()
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

        private void AddNipple()
        {
            cboC2Nipple.Items.Clear();
            var lstFmlNipple = new FilteredElementCollector(Global.UIDoc.Document).OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .Where(x => x.FamilyName.Contains("Nipple"))
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
                    AddDiameter(CurrentSegment);
            }
        }

        private void SprinkerUpForm_Load(object sender, EventArgs e)
        {
            AddFamilyType();
            AddNipple();

            rdnC2Tee.Checked = true;
            chkC2ConnectTee.Checked = false;
            chkC2Nipple.Enabled = false;
            cboC2Nipple.Enabled = false;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            PressCancel();
            App.m_SprinkerUpForm = null;

            this.Close();
        }

        private void btnApply_Click(object sender, EventArgs e)
        {
            isConnectTee = false;
            isConnectNipple = false;
            //fmlNipple = null;
            if (PipeSize == double.MaxValue)
                return;

            isD15 = rdnC2SprinklerD15.Checked;
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

        private void rdnC2Elbow_CheckedChanged(object sender, EventArgs e)
        {
            if (rdnC2Elbow.Checked)
                chkC2ConnectTee.Enabled = false;
            else
                chkC2ConnectTee.Enabled = true;
        }
    }
}