using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TotalMEPProject.UI.TotalMEPUI
{
    public partial class VerticalConnectionForm : Form
    {
        public bool IsType1;

        public VerticalConnectionForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            IsType1 = rbType1.Checked;
        }
    }
}