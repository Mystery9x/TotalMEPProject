using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
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

namespace TotalMEPProject.UI
{
    public partial class FastVerticalForm : System.Windows.Forms.Form
    {
        private Request.Request m_request = null;
        private RequestHandler m_handler = null;
        private ExternalEvent m_exEvent = null;

        private Dictionary<string, ElementId> _Levels = new Dictionary<string, ElementId>();

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

        public bool Siphon
        {
            get
            {
                return radSiPhon.Checked;
            }
        }

        public bool ByLevel
        {
            get
            {
                return true;
            }
        }

        public ElementId LevelId
        {
            get
            {
                if (cboLevel.SelectedItem == null)
                    return ElementId.InvalidElementId;

                string levelName = cboLevel.SelectedItem as string;
                if (_Levels.ContainsKey(levelName) == false)
                    return ElementId.InvalidElementId;

                return _Levels[levelName];
            }
        }

        public ElementId SiphonId
        {
            get
            {
                if (cboSiphon.SelectedItem == null)
                    return ElementId.InvalidElementId;

                return (cboSiphon.SelectedItem as ObjectItem).ObjectId;
            }
        }

        public double OffSet
        {
            get
            {
                double h = 0;
                if (double.TryParse(txtOffset.Text.Trim(), out h) == false)
                    return h;

                return h;
            }
        }

        public Type GetMEPType
        {
            get
            {
                if (cboMEPType.SelectedIndex == 0)
                    return typeof(Pipe);
                else if (cboMEPType.SelectedIndex == 1)
                    return typeof(Duct);
                else if (cboMEPType.SelectedIndex == 2)
                    return typeof(CableTray);
                else if (cboMEPType.SelectedIndex == 3)
                    return typeof(Conduit);

                return null;
            }
        }

        public FastVerticalForm(Dictionary<string, ElementId> levels, ExternalEvent exEvent, RequestHandler handler)
        {
            InitializeComponent();

            m_exEvent = exEvent;
            m_handler = handler;
            _Levels = levels;
            cboLevel.Enabled = true;
            txtOffset.Enabled = true;
            radElbow90.Checked = true;
            cboSiphon.Enabled = false;
        }

        #region Method

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

        public void PressCancel(int count = 2)
        {
            IWin32Window _revit_window = new WindowHandle(ComponentManager.ApplicationWindow);

            for (int i = 0; i < count; i++)
            {
                Press.PostMessage(_revit_window.Handle, (uint)Press.KEYBOARD_MSG.WM_KEYDOWN, (uint)Keys.Escape, 0);
            }
        }

        private void AddMEPType()
        {
            cboMEPType.Items.Clear();

            cboMEPType.Items.Add("Pipe");
            cboMEPType.Items.Add("Duct");
            cboMEPType.Items.Add("Cable Tray");
            cboMEPType.Items.Add("Conduit");

            AppUtils.ff(cboMEPType);

            if (cboMEPType.SelectedItem == null)
                cboMEPType.SelectedIndex = 0;
        }

        private void AddSiphon()
        {
            cboSiphon.Items.Clear();

            //Find intersection wit fitting
            var fittingBuilt = new ElementId(BuiltInCategory.OST_PipeFitting);
            FilteredElementCollector collector = new FilteredElementCollector(Global.UIDoc.Document);
            collector.OfClass(typeof(FamilySymbol));
            collector.OfCategoryId(fittingBuilt);

            if (collector.GetElementCount() == 0)
                return;

            var elements = collector.ToElements();
            if (elements.Count != 0)
            {
                foreach (FamilySymbol fitting in elements)
                {
                    ObjectItem item = new ObjectItem(fitting.Family.Name + " : " + fitting.Name, fitting.Id);

                    cboSiphon.Items.Add(item);
                }

                AppUtils.ff(cboSiphon);

                if (cboSiphon.SelectedItem == null)
                {
                    cboSiphon.SelectedIndex = 0;
                }
            }
        }

        #endregion Method

        #region Event

        private void FastVerticalForm_Load(object sender, EventArgs e)
        {
            AddMEPType();
            AddSiphon();

            cboLevel.Items.Clear();

            foreach (KeyValuePair<string, ElementId> keyPair in _Levels)
            {
                cboLevel.Items.Add(keyPair.Key);
            }

            AppUtils.ff(cboLevel);

            if (cboLevel.Items.Count != 0 && cboLevel.SelectedItem == null)
                cboLevel.SelectedIndex = 0;

            AppUtils.ff(txtOffset);

            AppUtils.ff(radElbow90);
            AppUtils.ff(radElbow45);
            AppUtils.ff(radSiPhon);
        }

        private void cboMEPType_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (GetMEPType == typeof(Pipe))
            {
                radElbow45.Enabled = true;
                radElbow90.Enabled = true;
            }
            else
            {
                radElbow45.Enabled = false;
                radElbow90.Enabled = false;

                radElbow90.Checked = true;
            }
        }

        private void txtOffset_KeyPress(object sender, KeyPressEventArgs e)
        {
            Common.NumberCheck(sender, e, true);
        }

        private void radSiPhon_CheckedChanged(object sender, EventArgs e)
        {
            cboSiphon.Enabled = radSiPhon.Checked;
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            if (LevelId == ElementId.InvalidElementId)
            {
                return;
            }
            if (SiphonId == ElementId.InvalidElementId)
            {
                return;
            }

            AppUtils.sa(cboLevel);
            AppUtils.sa(cboMEPType);

            AppUtils.sa(txtOffset);

            AppUtils.sa(radElbow90);
            AppUtils.sa(radElbow45);
            AppUtils.sa(radSiPhon);

            AppUtils.sa(cboSiphon);

            AppUtils.sa(this);
            SetFocus();
            MakeRequest(RequestId.FastVertical);
        }

        #endregion Event

        private void btnCancel_Click(object sender, EventArgs e)
        {
            PressCancel();

            AppUtils.sa(cboLevel);
            AppUtils.sa(cboMEPType);

            AppUtils.sa(txtOffset);

            AppUtils.sa(radElbow90);
            AppUtils.sa(radElbow45);
            AppUtils.sa(radSiPhon);

            AppUtils.sa(cboSiphon);

            AppUtils.sa(this);

            App.fastVerticalForm = null;

            this.Close();
        }
    }
}