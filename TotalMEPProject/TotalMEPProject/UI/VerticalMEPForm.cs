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
using TotalMEPProject.Request;

namespace TotalMEPProject.UI
{
    public partial class VerticalMEPForm : Form
    {
        #region Variable

        private RequestHandler m_handler = null;
        private ExternalEvent m_exEvent = null;

        #endregion Variable

        #region Constructor

        public VerticalMEPForm(ExternalEvent exEvent, RequestHandler handler)
        {
            InitializeComponent();
            m_handler = handler;
            m_exEvent = exEvent;
        }

        #endregion Constructor

        #region Method

        private void btnApply_Click(object sender, EventArgs e)
        {
        }

        #endregion Method
    }
}