﻿using Autodesk.Revit.DB;
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
        }

        private void _2LevelSmartForm_Load(object sender, EventArgs e)
        {
            AddFamilyType();
            AddNipple();
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
                .Where(x => x.FamilyName.Contains("Nipple"))
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

            this.Close();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (PipeSize == double.MaxValue)
                return;

            AppUtils.sa(cboC1FamilyType);
            AppUtils.sa(cboC1PipeSize);
            AppUtils.sa(cboC1NippleFamily);
            SetFocus();
            MakeRequest(RequestId.TwoLevelSmart_OK);
        }

        private void ckbNippleCreating_CheckedChanged(object sender, EventArgs e)
        {
            if (ckbNippleCreating.Checked)
                ckbC1ElbowConnection.Enabled = false;
            else
                ckbC1ElbowConnection.Enabled = true;
        }

        private void ckbC1ElbowConnection_CheckedChanged(object sender, EventArgs e)
        {
            if (ckbC1ElbowConnection.Checked)
            {
                ckbNippleCreating.Enabled = false;
                cboC1PipeSize.Enabled = false;
                cboC1NippleFamily.Enabled = false;
            }
            else
            {
                ckbNippleCreating.Enabled = true;
                cboC1PipeSize.Enabled = true;
                cboC1NippleFamily.Enabled = true;
            }
        }

        #endregion Event
    }
}