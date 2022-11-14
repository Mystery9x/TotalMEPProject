using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TotalMEPProject.Ultis;

namespace TotalMEPProject.UI
{
    public partial class HolyUpDownForm : Form
    {
        public MEPType m_MEPCurrent = MEPType.Pipe;

        public RunMode m_RunMode = RunMode.Apply;

        private Request.Request m_request;

        private Request.RequestHandler m_Handler;
        private ExternalEvent m_ExEvent;

        public HolyUpDownForm(ExternalEvent exEvent, Request.RequestHandler handler)
        {
            InitializeComponent();

            m_Handler = handler;
            m_ExEvent = exEvent;
        }

        public HolyUpDownForm()
        {
            InitializeComponent();
        }

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
                return radCustom.Checked;
            }
        }

        public bool ElbowCustom
        {
            get
            {
                return radCustom.Checked;
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

        private void btnApply_Click(object sender, EventArgs e)
        {
        }
    }

    public enum RunMode
    {
        Apply = 0,
        OK,
    }
}