﻿using Autodesk.Revit.UI;
using Autodesk.Windows;
using System;
using System.Windows.Forms;
using TotalMEPProject.Request;
using TotalMEPProject.Ultis;

namespace TotalMEPProject.UI
{
    public partial class HolyUpDownForm : Form
    {
        #region Variable

        public MEPType m_MEPCurrent = MEPType.Pipe;

        public RunMode m_runMode = RunMode.Apply;

        private RequestId m_request;

        private RequestHandler m_handler;

        private ExternalEvent m_exEvent;

        #endregion Variable

        #region Properties

        public bool Elbow90
        {
            get
            {
                return radElbow90.Checked;
            }
        }

        public bool Elbow45
        {
            get
            {
                return radElbow45.Checked;
            }
        }

        public bool ElbowCustom
        {
            get
            {
                return radCustom.Checked;
            }
        }

        public bool NotApply
        {
            get
            {
                return radNotApply.Checked;
            }
        }

        public double AngleCustom
        {
            get
            {
                double dvalue = 0;
                if (double.TryParse(txtAngle.Text.Trim(), out dvalue) == false)
                    return double.MinValue;
                else
                    return dvalue;
            }
        }

        public double Distance
        {
            get
            {
                double dvalue = 0;
                if (double.TryParse(txtDistance.Text.Trim(), out dvalue) == false)
                    return double.MinValue;
                else
                    return dvalue;
            }
        }

        public double UpStepValue
        {
            get
            {
                double dvalue = 0;
                if (double.TryParse(txtUpdownStepValue.Text.Trim(), out dvalue) == false)
                    return double.MinValue;
                else
                    return dvalue;
            }
        }

        public double UpElbowStepValue
        {
            get
            {
                double dvalue = 0;
                if (double.TryParse(txtEblowControlValue.Text.Trim(), out dvalue) == false)
                    return double.MinValue;
                else
                    return dvalue;
            }
        }

        public double DownStepValue
        {
            get
            {
                double dvalue = 0;
                if (double.TryParse(txtUpdownStepValue.Text.Trim(), out dvalue) == false)
                    return double.MinValue;
                else
                    return dvalue;
            }
        }

        public double UpElbowControlValue
        {
            get
            {
                double dvalue = 0;
                if (double.TryParse(txtEblowControlValue.Text.Trim(), out dvalue) == false)
                    return double.MinValue;
                else
                    return dvalue;
            }
        }

        public double DownElbowControlValue
        {
            get
            {
                double dvalue = 0;
                if (double.TryParse(txtEblowControlValue.Text.Trim(), out dvalue) == false)
                    return double.MinValue;
                else
                    return dvalue;
            }
        }

        #endregion Properties

        #region Constructor

        public HolyUpDownForm(ExternalEvent exEvent, RequestHandler handler)
        {
            InitializeComponent();

            m_handler = handler;
            m_exEvent = exEvent;
            txtAngle.Enabled = radCustom.Checked;

            radElbow45.Checked = true;
            txtDistance.Text = "300";
            txtUpdownStepValue.Text = "100";
            txtEblowControlValue.Text = "100";
        }

        #endregion Constructor

        #region Event

        private void btnApply_Click(object sender, EventArgs e)
        {
            if (Distance == double.MinValue && NotApply == false)
                return;
            m_runMode = RunMode.Apply;

            AppUtils.sa(txtDistance);
            AppUtils.sa(radElbow90);
            AppUtils.sa(radElbow45);
            AppUtils.sa(radCustom);
            AppUtils.sa(radNotApply);
            AppUtils.sa(txtAngle);
            AppUtils.sa(txtEblowControlValue);
            AppUtils.sa(txtUpdownStepValue);

            MakeRequest(GetRequestId(m_runMode));
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            m_runMode = RunMode.OK;

            AppUtils.sa(txtDistance);
            AppUtils.sa(radElbow90);
            AppUtils.sa(radElbow45);
            AppUtils.sa(radCustom);
            AppUtils.sa(radNotApply);
            AppUtils.sa(txtAngle);
            AppUtils.sa(txtEblowControlValue);
            AppUtils.sa(txtUpdownStepValue);

            MakeRequest(GetRequestId(RunMode.OK));

            this.Close();
        }

        private void btnUpStep_Click(object sender, EventArgs e)
        {
            if (UpStepValue == double.MinValue)
                return;
            AppUtils.sa(txtUpdownStepValue);
            MakeRequest(RequestId.HolyUpDown_UpStep);
        }

        private void btnDownStep_Click(object sender, EventArgs e)
        {
            if (UpStepValue == double.MinValue)
                return;
            AppUtils.sa(txtUpdownStepValue);
            MakeRequest(RequestId.HolyUpDown_DownStep);
        }

        private void btnUpElbowControl_Click(object sender, EventArgs e)
        {
            if (UpElbowStepValue == double.MinValue)
                return;
            AppUtils.sa(txtEblowControlValue);
            MakeRequest(RequestId.HolyUpDown_UpElbowControl);
        }

        private void btnDownElbowControl_Click(object sender, EventArgs e)
        {
            if (UpElbowStepValue == double.MinValue)
                return;
            AppUtils.sa(txtEblowControlValue);
            MakeRequest(RequestId.HolyUpDown_DownElbowControl);
        }

        private void txtAngle_KeyPress(object sender, KeyPressEventArgs e)
        {
            NumberCheck(sender, e, false);
        }

        private void txtUpdownStepValue_KeyPress(object sender, KeyPressEventArgs e)
        {
            NumberCheck(sender, e, true);
        }

        private void txtEblowControlValue_KeyPress(object sender, KeyPressEventArgs e)
        {
            NumberCheck(sender, e, false);
        }

        private void radCustom_CheckedChanged(object sender, EventArgs e)
        {
            txtAngle.Enabled = radCustom.Checked;
        }

        private void radNotApply_CheckedChanged(object sender, EventArgs e)
        {
            label4.Enabled = !radNotApply.Checked;
            btnUpElbowControl.Enabled = !radNotApply.Checked;
            btnDownElbowControl.Enabled = !radNotApply.Checked;
            txtEblowControlValue.Enabled = !radNotApply.Checked;
            txtDistance.Enabled = !radNotApply.Checked;

            if (radNotApply.Checked == false)
            {
                App.isApply = false;
            }
        }

        private void txtDistance_KeyPress(object sender, KeyPressEventArgs e)
        {
            NumberCheck(sender, e, true);
        }

        private void HolyUpDownForm_Load(object sender, EventArgs e)
        {
            AppUtils.ff(txtDistance);
            AppUtils.ff(radElbow90);
            AppUtils.ff(radElbow45);
            AppUtils.ff(radCustom);
            AppUtils.ff(radNotApply);
            AppUtils.ff(txtAngle);
            AppUtils.ff(txtEblowControlValue);
            AppUtils.ff(txtUpdownStepValue);

            txtAngle.Enabled = radCustom.Checked;
            label4.Enabled = !radNotApply.Checked;
            btnUpElbowControl.Enabled = !radNotApply.Checked;
            btnDownElbowControl.Enabled = !radNotApply.Checked;
            txtEblowControlValue.Enabled = !radNotApply.Checked;
        }

        #endregion Event

        #region Method

        public void MakeRequest(RequestId request)
        {
            m_handler.Request.Make(request);
            m_exEvent.Raise();
        }

        private RequestId GetRequestId(RunMode mode)
        {
            if (mode == RunMode.Apply)
            {
                RequestId id = RequestId.HolyUpDown_PickObjects;
                return id;
            }
            else
            {
                RequestId id = RequestId.HolyUpDown_OK;

                return id;
            }
        }

        public void PressCancel(int count = 2)
        {
            IWin32Window _revit_window = new WindowHandle(ComponentManager.ApplicationWindow);

            for (int i = 0; i < count; i++)
            {
                Press.PostMessage(_revit_window.Handle, (uint)Press.KEYBOARD_MSG.WM_KEYDOWN, (uint)Keys.Escape, 0);
            }
        }

        public static void NumberCheck(object sender, KeyPressEventArgs e, bool allowNegativeValue = false) // < 0
        {
            if (allowNegativeValue == false)
            {
                if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
                {
                    e.Handled = true;
                }

                // only allow one decimal point
                if ((e.KeyChar == '.') && ((sender as System.Windows.Forms.TextBox).Text.IndexOf('.') > -1))
                {
                    e.Handled = true;
                }
            }
            else
            {
                if (!char.IsControl(e.KeyChar) && (!char.IsDigit(e.KeyChar)) && (e.KeyChar != '.') && (e.KeyChar != '-'))
                    e.Handled = true;

                // only allow one decimal point
                if (e.KeyChar == '.' && (sender as System.Windows.Forms.TextBox).Text.IndexOf('.') > -1)
                    e.Handled = true;

                // only allow minus sign at the beginning
                if (e.KeyChar == '-' && (sender as System.Windows.Forms.TextBox).Text.IndexOf('-') > -1)
                    e.Handled = true;
            }
        }

        #endregion Method

        private void btnPreview_Click(object sender, EventArgs e)
        {
            if (btnPreview.Text == "Preview >>")
            {
                this.Size = this.MaximumSize;
                btnPreview.Text = "Preview <<";
            }
            else
            {
                this.Size = this.MinimumSize;
                btnPreview.Text = "Preview >>";
            }
        }
    }

    public enum RunMode
    {
        Apply = 0,
        ApplyLeftRight = 1,
        OK,
    }
}