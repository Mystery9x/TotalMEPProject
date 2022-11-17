using System;
using System.Windows.Forms;
using TotalMEPProject.Ultis;

namespace TotalMEPProject.UI
{
    public partial class FrmLogin : Form
    {
        public FrmLogin()
        {
            InitializeComponent();
            rbtnAccount.Enabled = false;
            rbtnLiscense.Checked = true;
        }

        private void rbtnLiscense_CheckedChanged(object sender, EventArgs e)
        {
            lblUserServer.Text = "Server Address:";
            lblPassword.Visible = !rbtnLiscense.Checked;
            tbPassword.Visible = !rbtnLiscense.Checked;
            linkForgot.Visible = !rbtnLiscense.Checked;
        }

        private void rbtnAccount_CheckedChanged(object sender, EventArgs e)
        {
            lblUserServer.Text = "UserName/Email:";
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
        }

        private void btnActive_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(tbUserServer.Text))
            {
                IO.ShowError("Please enter key to activate!", "TotalMEP");
                return;
            }

            var hardwareId = DataPC.Value();

            string errMess = "";

            bool success = LicenseUtils.RegistLicense(tbUserServer.Text, hardwareId, ref errMess);

            if (errMess != "")
            {
                IO.ShowError(errMess, "TotalMEP");
                return;
            }

            if (success)
            {
                if (App.cachedUiCtrApp != null)
                {
                    AppUtils.EnableItemRibbonTotalMEP(App.cachedUiCtrApp);
                    IO.ShowInfor("Active successfully!", "TotalMEP");
                }
            }

            this.DialogResult = DialogResult.OK;
        }
    }
}