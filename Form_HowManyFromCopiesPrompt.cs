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
    public partial class Form_HowManyFromCopiesPrompt : Form
    {

        public static int HowMany;
        


        public Form_HowManyFromCopiesPrompt()
        {
            InitializeComponent();

            textBox1.TabIndex = 0;
            button1.TabIndex = 1;
            textBox1.Select();
        }

        private void button1_Click(object sender, EventArgs e)//OK
        {
            HowMany = Int32.Parse(textBox1.Text) ;

            Form_AddImplementation form = new Form_AddImplementation();
            form.GetCopiesMethod();


            this.Close();
        }
    }
}
