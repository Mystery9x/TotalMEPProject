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
    public partial class FlexSprinklerForm : Form
    {
        #region Variable

        private Request.RequestHandler m_handler;

        private ExternalEvent m_exEvent;

        #endregion Variable

        #region Properties

        public bool IsCheckedD15
        {
            get => rbC4D15.Checked;
        }

        public bool IsCheckedD20
        {
            get => !rbC4D15.Checked;
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

        public double VerticalPipeLength
        {
            get
            {
                double dHeight = 0;

                if (double.TryParse(tbC4Length.Text.Trim(), out dHeight) == true)
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

        #endregion Method

        #region Event

        private void FlexSprinklerForm_Load(object sender, EventArgs e)
        {
            rbC4D15.Checked = true;
            rbC4Type1.Checked = true;
            AppUtils.ff(rbC4D15);
            AppUtils.ff(rbC4D20);
            AppUtils.ff(rbC4Type1);
            AppUtils.ff(rbC4Type2);
            AppUtils.ff(rbC4Type3);
            AppUtils.ff(tbC4Length);
        }

        private void btnC4Run_Click(object sender, EventArgs e)
        {
            if (VerticalPipeLength == double.MinValue)
                return;

            AppUtils.sa(rbC4D15);
            AppUtils.sa(rbC4D20);
            AppUtils.sa(rbC4Type1);
            AppUtils.sa(rbC4Type2);
            AppUtils.sa(rbC4Type3);
            AppUtils.sa(tbC4Length);

            MakeRequest(RequestId.FlexSprinker_RUN);
        }

        private void btnC4OK_Click(object sender, EventArgs e)
        {
            PressCancel();
            AppUtils.sa(rbC4D15);
            AppUtils.sa(rbC4D20);
            AppUtils.sa(rbC4Type1);
            AppUtils.sa(rbC4Type2);
            AppUtils.sa(rbC4Type3);
            AppUtils.sa(tbC4Length);
            App.m_flexSprinklerForm = null;
            this.Close();
        }

        private void tbC4Length_KeyPress(object sender, KeyPressEventArgs e)
        {
            Common.NumberCheck(sender, e, false);
        }

        #endregion Event
    }
}