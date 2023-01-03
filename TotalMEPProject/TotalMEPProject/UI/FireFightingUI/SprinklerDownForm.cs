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
    public partial class SprinklerDownForm : System.Windows.Forms.Form
    {
        private Request.Request m_request;

        private Request.RequestHandler m_handler;

        private ExternalEvent m_exEvent;
        public bool isD15 = false;

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

        public SprinklerDownForm(ExternalEvent exEvent, Request.RequestHandler handler)
        {
            InitializeComponent();

            m_handler = handler;
            m_exEvent = exEvent;
        }

        private void AddFamilyType()
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

        private void AddDiameter(Segment segment)
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
                    AddDiameter(CurrentSegment);
            }
        }

        private void SprinklerDownForm_Load(object sender, EventArgs e)
        {
            rdnC3Type1.Checked = true;
            AppUtils.ff(txbC3Length, "200");
            AddFamilyType();
        }

        private void btnC3Run_Click(object sender, EventArgs e)
        {
            if (Height_ == double.MinValue)
                return;
            isD15 = rdnC3D15.Checked;

            AppUtils.sa(cboC3PipeType);
            AppUtils.sa(cboC3PipeSize);
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

        private void txbC3Length_KeyPress(object sender, KeyPressEventArgs e)
        {
            Common.NumberCheck(sender, e, false);
        }

        private void btnC3OK_Click(object sender, EventArgs e)
        {
            PressCancel();
            App.m_SprinklerDownForm = null;

            this.Close();
        }
    }
}