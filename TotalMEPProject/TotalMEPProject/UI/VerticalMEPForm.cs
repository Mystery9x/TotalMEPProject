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
using Control = System.Windows.Forms.Control;

namespace TotalMEPProject.UI
{
    public partial class VerticalMEPForm : System.Windows.Forms.Form
    {
        #region Variable

        private MEPType m_MEPCurrent = MEPType.Pipe;
        private List<ElementId> FamilyTypes = new List<ElementId>();
        private List<ElementId> MEPSystemTypes = new List<ElementId>();
        private Segment CurrentSegment = null;
        private Dictionary<string, object> DictionaryMEPSizes = new Dictionary<string, object>();

        private List<Control> groupBoxDiameter = new List<Control>();
        private List<Control> groupBoxSystemType = new List<Control>();
        private List<Control> groupBoxWidht_Height = new List<Control>();
        private List<Control> groupBoxServiceType = new List<Control>();

        public bool IsShowDiameter = true;
        public bool IsShowSystemType = true;
        public bool IsShowWidht_Height = true;
        public bool IsShowServiceType = true;

        private Request.Request m_request = null;
        private RequestHandler m_handler = null;
        private ExternalEvent m_exEvent = null;

        #region Properties

        public Request.Request Request
        {
            get
            {
                return m_request;
            }
            private set
            {
                m_request = value;
            }
        }

        public ElementId LevelTopId
        {
            get
            {
                if (cboLevelTop.SelectedItem == null)
                    return ElementId.InvalidElementId;
                return (cboLevelTop.SelectedItem as ObjectItem).ObjectId;
            }
        }

        public ElementId LevelBottomId
        {
            get
            {
                if (cboLevelBottom.SelectedItem == null)
                    return ElementId.InvalidElementId;
                return (cboLevelBottom.SelectedItem as ObjectItem).ObjectId;
            }
        }

        public double OffsetTop
        {
            get
            {
                double value = 0;
                if (double.TryParse(txtOffsetTop.Text.Trim(), out value) == false)
                {
                    return double.MinValue;
                }

                return value;
            }
        }

        public double OffsetBottom
        {
            get
            {
                double value = 0;
                if (double.TryParse(txtOffsetBottom.Text.Trim(), out value) == false)
                {
                    return double.MinValue;
                }

                return value;
            }
        }

        public MEPType MEPType_
        {
            get
            {
                var enumType = (MEPType)Enum.Parse(typeof(MEPType), cboMEPObjects.SelectedItem.ToString());
                return enumType;
            }
        }

        public ElementId FamilyType
        {
            get
            {
                return (cboFamilyType.SelectedItem as ObjectItem).ObjectId;
            }
        }

        public ElementId SystemType
        {
            get
            {
                return (cboSystemType.SelectedItem as ObjectItem).ObjectId;
            }
        }

        public string ServiceType
        {
            get
            {
                return txtServiceType.Text.Trim();
            }
        }

        public double MEP_Width
        {
            get
            {
                var value = cboWidth.Text/*SelectedText*/.ToString().Trim();

                value = value.Replace(" mm", "");

                double dvalue = 0;
                if (value != string.Empty && double.TryParse(value, out dvalue) == false)
                {
                    return double.MaxValue;
                }

                return dvalue;
            }
        }

        public double MEP_Height
        {
            get
            {
                var value = cboHeight.Text.ToString().Trim();
                value = value.Replace(" mm", "");

                double dvalue = 0;
                if (value != string.Empty && double.TryParse(value, out dvalue) == false)
                {
                    return double.MaxValue;
                }

                return dvalue;
            }
        }

        public object MEPSize_
        {
            get
            {
                var value = cboDiameter.SelectedItem.ToString();

                if (DictionaryMEPSizes.ContainsKey(value) == true)
                    return DictionaryMEPSizes[value];
                return null;
            }
        }

        #endregion Properties

        #endregion Variable

        #region Constructor

        public VerticalMEPForm(ExternalEvent exEvent, RequestHandler handler)
        {
            InitializeComponent();
            m_handler = handler;
            m_exEvent = exEvent;
            groupBoxDiameter = new List<Control>() { lblDiameter, cboDiameter };
            groupBoxSystemType = new List<Control>() { lblSystemType, cboSystemType };
            groupBoxWidht_Height = new List<Control>() { lblWidth, cboWidth, lblHeight, cboHeight };
            groupBoxServiceType = new List<Control>() { lblServiceType, txtServiceType };
        }

        #endregion Constructor

        #region Method

        #region Event

        private void btnApply_Click(object sender, EventArgs e)
        {
            SaveControl();

            Run();
        }

        private void VerticalMEPForm_Load(object sender, EventArgs e)
        {
            AddMEPType();
            AddTopLevel();
            AddBottomLevel();

            AppUtils.ff(txtOffsetBottom);
            AppUtils.ff(txtOffsetTop);
            AppUtils.ff(cboDiameter);

            AppUtils.ff(this);
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            PressCancel();
            SaveControl();

            App.verticalMEPForm = null;

            this.Close();
        }

        public void PressCancel(int count = 2)
        {
            IWin32Window _revit_window = new WindowHandle(ComponentManager.ApplicationWindow);

            for (int i = 0; i < count; i++)
            {
                Press.PostMessage(_revit_window.Handle, (uint)Press.KEYBOARD_MSG.WM_KEYDOWN, (uint)Keys.Escape, 0);
            }
        }

        private void txtOffsetTop_KeyPress(object sender, KeyPressEventArgs e)
        {
            NumberCheck(sender, e, true);
        }

        public void NumberCheck(object sender, KeyPressEventArgs e, bool allowNegativeValue = false) // < 0
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

        private void txtOffsetBottom_KeyPress(object sender, KeyPressEventArgs e)
        {
            NumberCheck(sender, e, true);
        }

        #endregion Event

        #region Process

        public void MakeRequest(RequestId request)
        {
            m_handler.Request.Make(request);
            m_exEvent.Raise();
        }

        private void ShowHide(List<Control> list, bool showHide)
        {
            for (int i = 0; i < list.Count; i++)
            {
                var control = list[i];
                control.Visible = showHide;
            }
        }

        public void SetPositionControls()
        {
            ShowHide(groupBoxDiameter, IsShowDiameter);

            ShowHide(groupBoxServiceType, IsShowServiceType);

            ShowHide(groupBoxWidht_Height, IsShowWidht_Height);

            ShowHide(groupBoxSystemType, IsShowSystemType);
        }

        private void AddMEPType()
        {
            cboMEPObjects.Items.Clear();

            cboMEPObjects.Items.Add(MEPType.Pipe.ToString());
            cboMEPObjects.Items.Add(MEPType.Rectangular_Duct.ToString());
            cboMEPObjects.Items.Add(MEPType.Round_Duct.ToString());

            cboMEPObjects.Items.Add(MEPType.CableTray.ToString());
            cboMEPObjects.Items.Add(MEPType.Conduit.ToString());

            AppUtils.ff(cboMEPObjects, null, null);

            if (cboMEPObjects.SelectedItem == null && cboMEPObjects.Items.Count != 0)
                cboMEPObjects.SelectedIndex = 0;

            //Get current
            var mepType = cboMEPObjects.SelectedItem.ToString();
            m_MEPCurrent = (MEPType)Enum.Parse(typeof(MEPType), mepType);
        }

        private void AddTopLevel()
        {
            cboLevelTop.Items.Clear();

            var coll = new FilteredElementCollector(Global.UIDoc.Document).OfClass(typeof(Level));

            if (coll.GetElementCount() == 0)
                return;

            foreach (Level level in coll.ToElements())
            {
                ObjectItem item = new ObjectItem(level.Name, level.Id);
                cboLevelTop.Items.Add(item);
            }

            AppUtils.ff(cboLevelTop);

            if (cboLevelTop.SelectedItem == null && cboLevelTop.Items.Count != 0)
                cboLevelTop.SelectedIndex = 0;
        }

        private void AddBottomLevel()
        {
            cboLevelBottom.Items.Clear();

            var coll = new FilteredElementCollector(Global.UIDoc.Document).OfClass(typeof(Level));

            if (coll.GetElementCount() == 0)
                return;

            foreach (Level level in coll.ToElements())
            {
                ObjectItem item = new ObjectItem(level.Name, level.Id);
                cboLevelBottom.Items.Add(item);
            }

            AppUtils.ff(cboLevelBottom);

            if (cboLevelBottom.SelectedItem == null && cboLevelBottom.Items.Count != 0)
                cboLevelBottom.SelectedIndex = 0;
        }

        private void SaveControl()
        {
            AppUtils.sa(cboLevelBottom);
            AppUtils.sa(cboLevelTop);
            AppUtils.sa(cboMEPObjects);

            AppUtils.sa(txtOffsetBottom);
            AppUtils.sa(txtOffsetTop);
            AppUtils.sa(cboDiameter);
            AppUtils.sa(this);
        }

        public void Run()
        {
            SetFocus();

            MakeRequest(RequestId.VerticalMEP);
        }

        private void SetFocus()
        {
            IntPtr hBefore = DisplayService.GetForegroundWindow();
            DisplayService.SetForegroundWindow(ComponentManager.ApplicationWindow);
        }

        private void DisplayType()
        {
            AddFamilyType(m_MEPCurrent);

            AddSystemType(m_MEPCurrent);
        }

        private Type GetType(MEPType enumType)
        {
            if (enumType == MEPType.Pipe)
                return typeof(PipeType);
            else if (enumType == MEPType.CableTray)
                return typeof(CableTrayType);
            else if (enumType == MEPType.Oval_Duct || enumType == MEPType.Rectangular_Duct || enumType == MEPType.Round_Duct)
                return typeof(DuctType);
            else
                return typeof(ConduitType);
        }

        private void AddFamilyType(MEPType enumType)
        {
            cboFamilyType.Items.Clear();
            FilteredElementCollector pipeTypes = new FilteredElementCollector(Global.UIDoc.Document).OfClass(GetType(enumType));
            foreach (MEPCurveType type in pipeTypes)
            {
                if (enumType == MEPType.Oval_Duct || enumType == MEPType.Rectangular_Duct || enumType == MEPType.Round_Duct)
                {
                    DuctShape shape = DuctShape.Oval;

                    var familyName = type.LookupParameter("Family Name").AsString();

                    if (familyName.Contains("Rectangular"))
                    {
                        shape = DuctShape.Rectangular;
                    }
                    else if (familyName.Contains("Round"))
                    {
                        shape = DuctShape.Round;
                    }

                    if (enumType == MEPType.Oval_Duct && shape != DuctShape.Oval)
                        continue;

                    if (enumType == MEPType.Rectangular_Duct && shape != DuctShape.Rectangular)
                        continue;

                    if (enumType == MEPType.Round_Duct && shape != DuctShape.Round)
                        continue;
                }

                ObjectItem item = new ObjectItem(type.Name, type.Id);
                cboFamilyType.Items.Add(item);

                FamilyTypes.Add(type.Id);
            }

            AppUtils.ff(cboFamilyType, null, m_MEPCurrent.ToString());

            if (cboFamilyType.SelectedItem == null && cboFamilyType.Items.Count != 0)
                cboFamilyType.SelectedIndex = 0;
        }

        private Type GetSytemType(MEPType enumType)
        {
            if (enumType == MEPType.Pipe)
                return typeof(PipingSystemType);
            else if (enumType == MEPType.Oval_Duct || enumType == MEPType.Rectangular_Duct || enumType == MEPType.Round_Duct)
                return typeof(MechanicalSystemType);
            else
                return null;
        }

        private void AddSystemType(MEPType enumType)
        {
            cboSystemType.Items.Clear();

            var typeClass = GetSytemType(enumType);
            if (typeClass == null)
                return;

            FilteredElementCollector pipeTypes = new FilteredElementCollector(Global.UIDoc.Document).OfClass(typeClass);
            foreach (MEPSystemType type in pipeTypes)
            {
                ObjectItem item = new ObjectItem(type.Name, type.Id);
                cboSystemType.Items.Add(item);

                FamilyTypes.Add(type.Id);
            }

            AppUtils.ff(cboSystemType, null, m_MEPCurrent.ToString());

            if (cboSystemType.SelectedItem == null && cboSystemType.Items.Count != 0)
                cboSystemType.SelectedIndex = 0;
        }

        private void AddDuctSize(DuctSizeSettings settings, DuctShape shape)
        {
            foreach (KeyValuePair<DuctShape, DuctSizes> keyPair in settings)
            {
                if (keyPair.Key != shape)
                    continue;

                if (keyPair.Key == DuctShape.Round)
                {
                    cboDiameter.Items.Clear();
                    DictionaryMEPSizes.Clear();

                    foreach (MEPSize size in keyPair.Value)
                    {
                        var value = FeetToMmString(size.NominalDiameter) + " mm";

                        cboDiameter.Items.Add(value);

                        DictionaryMEPSizes.Add(value, size);
                    }

                    AppUtils.ff(cboDiameter, null, m_MEPCurrent.ToString());

                    if (cboDiameter.SelectedItem == null && cboDiameter.Items.Count != 0)
                        cboDiameter.SelectedIndex = 0;
                }
                else
                {
                    cboWidth.Items.Clear();
                    cboHeight.Items.Clear();

                    DictionaryMEPSizes.Clear();

                    foreach (MEPSize size in keyPair.Value)
                    {
                        var value = FeetToMmString(size.NominalDiameter) + " mm";

                        cboWidth.Items.Add(value);
                        cboHeight.Items.Add(value);
                        DictionaryMEPSizes.Add(value, size);
                    }

                    AppUtils.ff(cboWidth, null, m_MEPCurrent.ToString());
                    AppUtils.ff(cboHeight, null, m_MEPCurrent.ToString());

                    if (cboWidth.SelectedItem == null && cboWidth.Items.Count != 0)
                        cboWidth.SelectedIndex = 0;

                    if (cboHeight.SelectedItem == null && cboHeight.Items.Count != 0)
                        cboHeight.SelectedIndex = 0;
                }
            }
        }

        private void AddDiameter(Segment segment)
        {
            cboDiameter.Items.Clear();
            DictionaryMEPSizes.Clear();

            foreach (MEPSize size in segment.GetSizes())
            {
                var value = FeetToMmString(size.NominalDiameter) + " mm";

                cboDiameter.Items.Add(value);

                DictionaryMEPSizes.Add(value, size);
            }

            AppUtils.ff(cboDiameter, null, m_MEPCurrent.ToString());

            if (cboDiameter.SelectedItem == null && cboDiameter.Items.Count != 0)
                cboDiameter.SelectedIndex = 0;
        }

        private void AddCableTraySizes(CableTraySizes sizes)
        {
            cboWidth.Items.Clear();
            cboHeight.Items.Clear();

            DictionaryMEPSizes.Clear();

            foreach (MEPSize size in sizes)
            {
                var value = FeetToMmString(size.NominalDiameter) + " mm";

                cboWidth.Items.Add(value);
                cboHeight.Items.Add(value);
                DictionaryMEPSizes.Add(value, size);
            }

            AppUtils.ff(cboWidth, null, m_MEPCurrent.ToString());
            AppUtils.ff(cboHeight, null, m_MEPCurrent.ToString());

            if (cboWidth.SelectedItem == null && cboWidth.Items.Count != 0)
                cboWidth.SelectedIndex = 0;

            if (cboHeight.SelectedItem == null && cboHeight.Items.Count != 0)
                cboHeight.SelectedIndex = 0;
        }

        private void AddDiameter(ConduitSizeSettings settings, string standardName)
        {
            cboDiameter.Items.Clear();
            DictionaryMEPSizes.Clear();

            foreach (KeyValuePair<string, ConduitSizes> keyPair in settings)
            {
                if (keyPair.Key != standardName)
                    continue;

                foreach (ConduitSize size in keyPair.Value)
                {
                    var value = FeetToMmString(size.NominalDiameter) + " mm";

                    cboDiameter.Items.Add(value);

                    DictionaryMEPSizes.Add(value, size);
                }
            }

            AppUtils.ff(cboDiameter, null, m_MEPCurrent.ToString());

            if (cboDiameter.SelectedItem == null && cboDiameter.Items.Count != 0)
                cboDiameter.SelectedIndex = 0;
        }

        public string FeetToMmString(double a)
        {
            return (a / Common.mmToFT).ToString("0.##");
        }

        #endregion Process

        #endregion Method
    }
}