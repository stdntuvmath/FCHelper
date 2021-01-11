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
    public partial class Form_ProcessFile : Form
    {
        public Form_ProcessFile()
        {
            InitializeComponent();
        }

        private void Form_ProcessFile_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)//back
        {
            if (webBrowser1.CanGoBack)
            {
                webBrowser1.GoBack();
            }
            
            //SimulateRickClickForContextMenuSelection rightclick = new SimulateRickClickForContextMenuSelection();
            //rightclick.RightClickMehtod();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog fbd = new FolderBrowserDialog() { Description = "Select Your Path" })
            {
                if (fbd.ShowDialog()==DialogResult.OK)
                {
                    webBrowser1.Url = new Uri(fbd.SelectedPath);
                    textBox1.Text = fbd.SelectedPath;
                }
            }

        }

        private void button2_Click(object sender, EventArgs e)//foward
        {
            if (webBrowser1.CanGoForward)
            {
                webBrowser1.GoForward();
            }
        }
    }
}
