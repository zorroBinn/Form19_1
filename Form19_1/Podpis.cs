using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TextBox;

namespace Form19_1
{
    public partial class Podpis : Form
    {
        private Main parentForm;

        public Podpis(Main parent)
        {
            InitializeComponent();
            this.parentForm = parent;
        }

        private void button_zapomn_Click(object sender, EventArgs e)
        {
            parentForm.UpdateButtonColor();
            this.Close();
        }
    }
}
