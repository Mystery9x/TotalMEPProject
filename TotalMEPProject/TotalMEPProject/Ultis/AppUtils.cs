using Microsoft.Win32;
using System;
using System.Windows.Forms;

namespace TotalMEPProject.Ultis
{
    public class AppUtils
    {
        public static void EnableItemRibbonTotalMEP(Autodesk.Revit.UI.UIControlledApplication app)
        {
            var ribbonPanels = app.GetRibbonPanels("TotalMEP");
            foreach (var item in ribbonPanels)
            {
                item.Enabled = true;
            }
        }

        public static string a(string name)
        {
            try
            {
                RegistryKey key = Registry.CurrentUser.OpenSubKey("Software\\MEPGenerator\\MEPGenerator\\RecentInput");
                if (key == null)
                {
                    return null;
                }
                else
                {
                    var obj = key.GetValue(name);
                    if (obj != null && key.GetValueKind(name) == RegistryValueKind.String)
                    {
                        return obj.ToString();
                    }
                }
                return null;
            }
            catch (Exception ex)
            {
                //Logger.Log.ErrorFormat("Không đọc được registry {0}", ex.Message);
                return null;
            }
        }

        public static bool f(string name, string value)
        {
            try
            {
                RegistryKey key = Registry.CurrentUser.CreateSubKey("Software\\MEPGenerator\\MEPGenerator\\RecentInput");
                if (key == null)
                {
                    return false;
                }
                else
                {
                    key.SetValue(name, value);
                    return true;
                }
            }
            catch (Exception ex)
            {
                //Logger.Log.ErrorFormat("Không ghi được registry {0}", ex.Message);
                return false;
            }
        }

        public static void ff(Control ctrl, string def = null, string group = null)
        {
            string key = string.Format("${0}${1}", ctrl.TopLevelControl/*Parent*/.Name, ctrl.Name);
            if (group != null && group != string.Empty)
            {
                key = string.Format("${0}${1}${2}", ctrl.TopLevelControl/*Parent*/.Name, ctrl.Name, group);
            }

            string value = a(key);
            if (value != null)
            {
                if (ctrl is ComboBox)
                {
                    var combobox = ctrl as ComboBox;
                    if (combobox.Items.Count == 0)
                    {
                        //Add
                        if (key.Contains("$CreateMEPObjectsForm$cboOffsetPipe_Conduit"))
                        {
                            combobox.Items.Add(value);
                        }
                    }
                    int index = combobox.FindStringExact(value);
                    if (index != -1)
                        combobox.SelectedIndex = index;
                }
                else if (ctrl is RadioButton)
                {
                    if (value == "true" || value == "false")
                    {
                        var rad = ctrl as RadioButton;
                        rad.Checked = value == "true" ? true : false;
                    }
                }
                else if (ctrl is CheckBox)
                {
                    if (value == "true" || value == "false")
                    {
                        var rad = ctrl as CheckBox;
                        rad.Checked = value == "true" ? true : false;
                    }
                }
                else if (ctrl is Form)
                {
                    if (value != null)
                    {
                        value = value.Replace("X=", "");
                        value = value.Replace("Y=", "");

                        value = value.Replace("{", "");
                        value = value.Replace("}", "");

                        var find = value.IndexOf(",");

                        if (find != -1)
                        {
                            var szx = value.Split(',')[0];
                            var szy = value.Split(',')[1];

                            int x = 0;
                            int y = 0;

                            if (int.TryParse(szx, out x) == true && int.TryParse(szy, out y))
                            {
                                (ctrl as Form).Location = new System.Drawing.Point(x, y);
                            }
                        }
                    }
                }
                else
                    ctrl.Text = value;
            }
            else
            {
                if (ctrl is ComboBox || ctrl is RadioButton || ctrl is CheckBox)
                {
                    return;
                }
                if (string.IsNullOrEmpty(def) == false)
                    ctrl.Text = def;
            }
        }

        public static void sa(Control ctrl, string group = null)
        {
            string key = string.Format("${0}${1}", ctrl.TopLevelControl/*Parent*/.Name, ctrl.Name);

            if (group != null && group != string.Empty)
            {
                key = string.Format("${0}${1}${2}", ctrl.TopLevelControl/*Parent*/.Name, ctrl.Name, group);
            }

            string value = ctrl.Text;
            if (ctrl is RadioButton)
            {
                value = (ctrl as RadioButton).Checked == true ? "true" : "false";
            }
            else if (ctrl is CheckBox)
            {
                value = (ctrl as CheckBox).Checked == true ? "true" : "false";
            }
            else if (ctrl is Form)
            {
                value = (ctrl as Form).Location.ToString();
            }

            f(key, value);
        }
    }
}