using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FCHelper_v001
{
    public partial class Form_PFMContactRequest : Form
    {
        public Form_PFMContactRequest()
        {
            InitializeComponent();
        }

        private void Form_PFMContactRequest_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            button1.Visible = false;
            PushButtonsOnPFM_ContactChange push = new PushButtonsOnPFM_ContactChange();
            push.ActiveatePFMMethod();

        }
    }
}
