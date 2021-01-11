using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;

namespace FCHelper_v001
{
    public partial class Form_NewLogEntryInOpenImplementationNotes : Form
    {

        public static string TodaysDate;
        public static string ERID;
        public static string EntryType;
        public static string Regarding;


        public Form_NewLogEntryInOpenImplementationNotes()
        {
            
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)//Submit Button
        {
            
            


            TodaysDate = dataGridView1.Rows[0].Cells[0].Value.ToString();
            ERID = dataGridView1.Rows[0].Cells[1].Value.ToString();
            EntryType = dataGridView1.Rows[0].Cells[2].Value.ToString();
            Regarding = dataGridView1.Rows[0].Cells[3].Value.ToString();


            Form_OpenImplementationList form = new Form_OpenImplementationList();



            //CheckIfFormIsOpen checkForm = new CheckIfFormIsOpen();

            //bool formOpen = checkForm.CheckIfFormIsOpenMethod("Form_OpenImplementationList");
            //if (formOpen == true)
            //{
            //    Form_OpenImplementationList form_OpenImplementationList = new Form_OpenImplementationList();
            //    form_OpenImplementationList.Activate();
            //}
            //else
            //{
            //    MessageBox.Show("Didnt work.");
            //}
            form.GetNewLogEntryDataAndDisplay();
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)//Cancel Button
        {
            this.Dispose();
        }

        private void Form_NewLogEntryInOpenImplementationNotes_Load(object sender, EventArgs e)
        {
            string erid = Form_OpenImplementationList.EmployerID;
            DateTime today = DateTime.Today;
            string todaysDate = today.ToString("MM/dd/yyyy");


            dataGridView1.Rows[0].Cells[0].Value = todaysDate;
            dataGridView1.Rows[0].Cells[1].Value = erid;
        }
    }
}
