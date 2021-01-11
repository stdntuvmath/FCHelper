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
    public partial class Form_ClaimsTroubleshooting : Form
    {
        public Form_ClaimsTroubleshooting()
        {
            InitializeComponent();
        }

        private void ClaimsTroubleshootingForm_Load(object sender, EventArgs e)
        {

        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            int selectedIndex = checkedListBox1.SelectedIndex;


            if (checkedListBox1.GetItemChecked(0) == true)
            {
                string labelText = "Is the Claims feed a Legacy Feed or Connected Claims Feed?";
                Controls.Add(new Label { Location = new Point(13, 150), AutoSize = true, Text = labelText, Name = "index0Label" });

                Controls.Add(new ComboBox {Name = "combobox1",Location = new Point(300,150),Text = "Legacy", });
                
            }
            else if(checkedListBox1.GetItemChecked(0) == false)
            {
                Controls.Remove(new Label { });
            }
            
            /*
            if (checkedListBox1.GetItemCheckState(0)==CheckState.Checked)
            {
                string labelText = "Is the Claims feed a Legacy Feed or Connected Claims Feed?";
                Controls.Add(new Label {Location = new Point(13,150), AutoSize = true, Text = labelText });

            }
            */
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)//Conclusion Button
        {

        }
    }
}
