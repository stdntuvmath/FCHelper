using System;
using System.Windows;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Microsoft.VisualBasic.FileIO;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Data.OleDb;
using Microsoft.Office.Interop.Excel;
//using LinqToExcel;



namespace FCHelper_v001
{
    public partial class Form_AddImplementation : System.Windows.Forms.Form
    {
        System.Data.DataTable dataTable1 = new System.Data.DataTable();//makes dataTable available

        // public string[] allMyFields;//global variable
        // public string[] AllFields { get => allMyFields; set allMyFields => }//pulls Word file data out of AddImplementation_DragDrop_1 method
        // and shunts it to allFields global variable, above, making
        // it accessable to this whole class

        string userName = System.Environment.UserName;//gives windows username
        

        public Form_AddImplementation()
        {
            InitializeComponent();

           /*

            string fileName = "C:\\Users\\" + userName + "\\Documents\\ImpList.txt";
            
            FileInfo fi = new FileInfo(fileName);



            if (!fi.Exists)//if the file doesn't exist, creat a txt file and insert tab delimited headings into the text file
            {
                

                using (StreamWriter sw = fi.CreateText())
                {
                    sw.WriteLine("Employer Name\tEmployer ID\tRegion\tSegment\t" +
                        "Effective Date\tCurrent Products\tProducts Being Added\t" +
                        "New Implementation\tIM/AM\tImplementation Deadline\t" +
                        "File Type\tContacts Name\tContacts Email\tContacts Phone");
                }

                
            }

*/
            

        }



        private void AddImplementation_Load(object sender, EventArgs e)
        {
            this.Location = new System.Drawing.Point(60, 60);

            if (Form_OpenImplementationList.DuplicateForm == true)
            {
                this.button3.PerformClick();
                textBox1.Text = Form_OpenImplementationList.EmployerName;
                textBox2.Text = Form_OpenImplementationList.EmployerID;
                textBox3.Text = Form_OpenImplementationList.Region;
                textBox4.Text = Form_OpenImplementationList.Segment;
                textBox5.Text = Form_OpenImplementationList.EffectiveDate;
                textBox6.Text = Form_OpenImplementationList.CurrentProducts;
                textBox7.Text = Form_OpenImplementationList.AddingProduct;
                textBox8.Text = Form_OpenImplementationList.NewImp;
                textBox9.Text = Form_OpenImplementationList.AMIMInvolved;
                textBox10.Text = Form_OpenImplementationList.ImpDeadline;
                textBox11.Text = Form_OpenImplementationList.SFTPCreds;

                dataGridView1.Rows.Add();

                dataGridView1.Rows[0].Cells[0].Value = Form_OpenImplementationList.InternalContactName1;
                dataGridView1.Rows[0].Cells[1].Value = Form_OpenImplementationList.InternalContactPhone1;
                dataGridView1.Rows[0].Cells[2].Value = Form_OpenImplementationList.InternalContactEmail1;
                dataGridView1.Rows[0].Cells[3].Value = Form_OpenImplementationList.InternalContactType1;

                dataGridView1.Rows.Add();
        
                dataGridView1.Rows[1].Cells[0].Value = Form_OpenImplementationList.InternalContactName2;
                dataGridView1.Rows[1].Cells[1].Value = Form_OpenImplementationList.InternalContactPhone2;
                dataGridView1.Rows[1].Cells[2].Value = Form_OpenImplementationList.InternalContactEmail2;
                dataGridView1.Rows[1].Cells[3].Value = Form_OpenImplementationList.InternalContactType2;

                dataGridView1.Rows.Add();

                dataGridView1.Rows[2].Cells[0].Value = Form_OpenImplementationList.InternalContactName3;
                dataGridView1.Rows[2].Cells[1].Value = Form_OpenImplementationList.InternalContactPhone3;
                dataGridView1.Rows[2].Cells[2].Value = Form_OpenImplementationList.InternalContactEmail3;
                dataGridView1.Rows[2].Cells[3].Value = Form_OpenImplementationList.InternalContactType3;

                dataGridView1.Rows.Add();

                dataGridView1.Rows[3].Cells[0].Value = Form_OpenImplementationList.InternalContactName4;
                dataGridView1.Rows[3].Cells[1].Value = Form_OpenImplementationList.InternalContactPhone4;
                dataGridView1.Rows[3].Cells[2].Value = Form_OpenImplementationList.InternalContactEmail4;
                dataGridView1.Rows[3].Cells[3].Value = Form_OpenImplementationList.InternalContactType4;

                dataGridView2.Rows.Add();

                dataGridView2.Rows[0].Cells[0].Value = Form_OpenImplementationList.ExternalContactName1;
                dataGridView2.Rows[0].Cells[1].Value = Form_OpenImplementationList.ExternalContactPhone1;
                dataGridView2.Rows[0].Cells[2].Value = Form_OpenImplementationList.ExternalContactEmail1;
                dataGridView2.Rows[0].Cells[3].Value = Form_OpenImplementationList.ExternalContactType1;

                dataGridView2.Rows.Add();

                dataGridView2.Rows[1].Cells[0].Value = Form_OpenImplementationList.ExternalContactName2;
                dataGridView2.Rows[1].Cells[1].Value = Form_OpenImplementationList.ExternalContactPhone2;
                dataGridView2.Rows[1].Cells[2].Value = Form_OpenImplementationList.ExternalContactEmail2;
                dataGridView2.Rows[1].Cells[3].Value = Form_OpenImplementationList.ExternalContactType2;

                dataGridView2.Rows.Add();

                dataGridView2.Rows[2].Cells[0].Value = Form_OpenImplementationList.ExternalContactName3;
                dataGridView2.Rows[2].Cells[1].Value = Form_OpenImplementationList.ExternalContactPhone3;
                dataGridView2.Rows[2].Cells[2].Value = Form_OpenImplementationList.ExternalContactEmail3;
                dataGridView2.Rows[2].Cells[3].Value = Form_OpenImplementationList.ExternalContactType3;

                dataGridView2.Rows.Add();

                dataGridView2.Rows[3].Cells[0].Value = Form_OpenImplementationList.ExternalContactName4;
                dataGridView2.Rows[3].Cells[1].Value = Form_OpenImplementationList.ExternalContactPhone4;
                dataGridView2.Rows[3].Cells[2].Value = Form_OpenImplementationList.ExternalContactEmail4;
                dataGridView2.Rows[3].Cells[3].Value = Form_OpenImplementationList.ExternalContactType4;
            }   

        }

          

        private void AddImplementation_DragEnter_1(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;
        }

        public void AddImplementation_DragDrop_1(object sender, DragEventArgs e)
        {
            try
            {
                string[] fileConsultantRequestFormFile = (string[])e.Data.GetData(DataFormats.FileDrop, false);

                foreach (string file in fileConsultantRequestFormFile)//iterate through all files dropped into the form
                {

                    if (file.Contains("FileConsultantRequestForm") || file.Contains("File Consultant Request Form") || file.Contains("File Consultant Request") || file.Contains("FCR"))
                    {

                        string employerName,
                    employerID,
                    region,   segment,
                    benefitEffectiveDate,   currentProducts,
                     addedProducts,   newImpFlag,
                    IM_AM,   impDeadline,   sftpFlag,   contactName,
                      contactphoneNumber,   contactEmail,   contactType,
                      fileLayout;
                        GetWordFileData passInputToMethod = new GetWordFileData();//create instance variable

                        //pull word data off of the FileConsultantRequestForm.docx into this form
                        passInputToMethod.GetWordFileDataMethod(file, out employerName, out  employerID,
                                                                out  region, out  segment,
                                                                out  benefitEffectiveDate, out  currentProducts,
                                                                out  addedProducts, out  newImpFlag,
                                                                out  IM_AM, out  impDeadline, out  sftpFlag, out  contactName,
                                                                out  contactphoneNumber, out  contactEmail, out  contactType,
                                                                out  fileLayout);


                        Regex rg = new Regex("[A-Za-z0-9/ .]^");

                        bool thing = rg.IsMatch("•");
                        if (thing == true)
                        {
                            employerName = employerName + rg.Replace("•", "");
                            employerID = employerID + rg.Replace("•", "");
                        }

                        

                        //unhide the controls and put FileConsultantRequestForm.docx data on the controls of this form
                        this.Activate();
                        this.label12.Visible = false;
                        this.label1.Show();
                        this.label2.Show();
                        this.label3.Show();
                        this.label4.Show();
                        this.label5.Show();
                        this.label6.Show();
                        this.label7.Show();
                        this.label8.Show();
                        this.label9.Show();
                        this.label10.Show();
                        this.label11.Show();
                        this.label13.Show();

                        this.textBox1.Show();
                        this.textBox2.Show();
                        this.textBox3.Show();
                        this.textBox4.Show();
                        this.textBox5.Show();
                        this.textBox6.Show();
                        this.textBox7.Show();
                        this.textBox8.Show();
                        this.textBox9.Show();
                        this.textBox10.Show();
                        this.textBox11.Show();

                        this.dataGridView1.Show();
                        this.dataGridView2.Show();

                        this.button1.Show();
                        this.button2.Visible = false;
                        this.button3.Visible = false;
                        this.button4.Show();
                        this.button5.Show();

                        this.textBox1.Text = employerName;
                        this.textBox2.Text = employerID;
                        this.textBox3.Text = region;
                        this.textBox4.Text = segment;
                        this.textBox5.Text = benefitEffectiveDate;
                        this.textBox6.Text = currentProducts;
                        this.textBox7.Text = addedProducts;
                        this.textBox8.Text = newImpFlag;
                        this.textBox9.Text = IM_AM; 
                        this.textBox10.Text = impDeadline;
                        this.textBox11.Text = sftpFlag;


                        dataGridView1[0,0].Value = IM_AM;
                        string[] dragDropDataToGridView2 = { contactName, contactphoneNumber, contactEmail, contactType, fileLayout };
                        this.dataGridView2.Rows.Add(dragDropDataToGridView2);

                        //set the tab order for all textboxes on this form

                        this.textBox1.TabIndex = 0;
                        this.textBox2.TabIndex = 1;
                        this.textBox3.TabIndex = 2;
                        this.textBox4.TabIndex = 3;
                        this.textBox5.TabIndex = 4;
                        this.textBox6.TabIndex = 5;
                        this.textBox7.TabIndex = 6;
                        this.textBox8.TabIndex = 7;
                        this.textBox9.TabIndex = 8;
                        this.textBox10.TabIndex = 9;
                        this.textBox11.TabIndex = 10;
                        this.dataGridView1.TabIndex = 11;


                    }

                    else if (file.Contains("docx") || file.Contains("doc"))//if the file is not the FileConsultantRequestForm.docx message the user
                    {

                        string message = "Is this a FileConsultantRequestForm file?\n" +
                                         "If so, please rename the file to include " +
                                         "FileConsultantRequestForm.docx and re-drop the file. If not, " +
                                         "please select No and drag and drop the correct file.";

                        string caption = "Filename other than FileConsultantRequestForm.docx";

                        MessageBox.Show(message, caption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                    }

                   

                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Method: AddImplementation\r Something prevented data from pulling from the FileConsultantRequestForm.docx.\r\r"+ex);
            }
           
        }



       





        private void textBox1_TextChanged(object sender, EventArgs e)
        {
           
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
          
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
          
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
          
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
          
        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
          
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
           
        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {
           
        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {
           

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        public void SetTable(System.Data.DataTable table)
        {
           

            string[] impListRowArray = {textBox1.Text, textBox2.Text,
                                        textBox3.Text, textBox4.Text,
                                        textBox5.Text, textBox6.Text,
                                        textBox7.Text, textBox8.Text,
                                        textBox9.Text, textBox10.Text};

            
        }


        private void button1_Click(object sender, EventArgs e)//Add Implementation Button
        {
            //int tableID = 1;

            string ername  = textBox1.Text;
            string erid    = textBox2.Text;
            string region  = textBox3.Text;
            string segment = textBox4.Text;
            string effDate = textBox5.Text;
            string curProd = textBox6.Text;
            string addProd = textBox7.Text;
            string newImp  = textBox8.Text;
            string AM_IM   = textBox9.Text;
            string impDdline = textBox10.Text;
            string sftpFlag = textBox11.Text;

            //filter textbox1 and 2 to remove carriage returns

            textBox1.Text = textBox1.Text.TrimEnd('\r', '\n');
            textBox2.Text = textBox2.Text.TrimEnd('\r', '\n');

            //create employer directory
            Directory.CreateDirectory(@"C:\Users\14025\Documents\File Consultants\Brandon\" + textBox1.Text + "_" + textBox2.Text);


            if (dataGridView1.Rows.Count < 4)
            {
                dataGridView1.Rows.Add();
                dataGridView1.Rows.Add();
                dataGridView1.Rows.Add();
            }

            if (dataGridView2.Rows.Count < 4)
            {
                dataGridView2.Rows.Add();
                dataGridView2.Rows.Add();
                dataGridView2.Rows.Add();
            }

            string inConName = (string)dataGridView1[0, 0].Value; 
            string inConPhone = (string)dataGridView1[1, 0].Value;
            string inConEmail = (string)dataGridView1[2, 0].Value;
            string inConType = (string)dataGridView1[3, 0].Value;            

            string exConName = (string)dataGridView2[0,0].Value; 
            string exConPhone = (string)dataGridView2[1, 0].Value;
            string exConEmail = (string)dataGridView2[2, 0].Value;
            string exConType = (string)dataGridView2[3, 0].Value;
            string fileType = (string)dataGridView2[4, 0].Value;

            string inConName2 = (string)dataGridView1[0, 1].Value;
            string inConPhone2 = (string)dataGridView1[1, 1].Value;
            string inConEmail2 = (string)dataGridView1[2, 1].Value;
            string inConType2 = (string)dataGridView1[3, 1].Value;

            string inConName3 = (string)dataGridView1[0, 2].Value;
            string inConPhone3 = (string)dataGridView1[1, 2].Value;
            string inConEmail3 = (string)dataGridView1[2, 2].Value;
            string inConType3 = (string)dataGridView1[3, 2].Value;

            string inConName4 = (string)dataGridView1[0, 3].Value;
            string inConPhone4 = (string)dataGridView1[1, 3].Value;
            string inConEmail4 = (string)dataGridView1[2, 3].Value;
            string inConType4 = (string)dataGridView1[3, 3].Value;

            string exConName2 = (string)dataGridView2[0, 1].Value;
            string exConPhone2 = (string)dataGridView2[1, 1].Value;
            string exConEmail2 = (string)dataGridView2[2, 1].Value;
            string exConType2 = (string)dataGridView2[3, 1].Value;

            string exConName3 = (string)dataGridView2[0, 2].Value;
            string exConPhone3 = (string)dataGridView2[1, 2].Value;
            string exConEmail3 = (string)dataGridView2[2, 2].Value;
            string exConType3 = (string)dataGridView2[3, 2].Value;

            string exConName4 = (string)dataGridView2[0, 3].Value;
            string exConPhone4 = (string)dataGridView2[1, 3].Value;
            string exConEmail4 = (string)dataGridView2[2, 3].Value;
            string exConType4 = (string)dataGridView2[3, 3].Value;


            //string stuff = tableID++.ToString();

            ExcelDataBasePush db = new ExcelDataBasePush();
            db.ExcelDataBasePushMethod(ername, erid, region, segment,
                                    effDate, curProd, addProd, newImp,
                                    AM_IM, impDdline, sftpFlag, 
                                    inConName, inConPhone, inConEmail, inConType, 
                                    exConName, exConPhone, exConEmail, exConType, fileType,
                                    inConName2, inConPhone2, inConEmail2, inConType2,
                                    inConName3, inConPhone3, inConEmail3, inConType3,
                                    inConName4, inConPhone4, inConEmail4, inConType4,
                                    exConName2, exConPhone2, exConEmail2, exConType2,
                                    exConName3, exConPhone3, exConEmail3, exConType3,
                                    exConName4, exConPhone4, exConEmail4, exConType4,
                                    string.Empty, string.Empty, string.Empty, string.Empty,
                                                string.Empty, string.Empty, string.Empty, string.Empty,
                                                string.Empty, string.Empty);
            SendKeys.Send("{ENTER}");
            try
            {
                //find non-local instance of another form and close it
                for (int index = System.Windows.Forms.Application.OpenForms.Count - 1; index >= 0; index--)
                {
                    if (System.Windows.Forms.Application.OpenForms[1].Name == "Form_OpenImplementationList")
                    {
                        System.Windows.Forms.Application.OpenForms[1].Close();
                    }
                }

                Form_OpenImplementationList form_OpenImplementationList = new Form_OpenImplementationList();
                form_OpenImplementationList.Show();

                this.Close();
            }
            catch (Exception ex)
            {
                
            }
            



        }

        private void button2_Click(object sender, EventArgs e)//Cancel
        {
            this.Close();
            
        }

        private void button3_Click(object sender, EventArgs e)//Manual Form
        {
            //unhide the controls and put FileConsultantRequestForm.docx data on the controls of this form
            this.Activate();
            this.label12.Visible = false;
            this.label1.Show();
            this.label2.Show();
            this.label3.Show();
            this.label4.Show();
            this.label5.Show();
            this.label6.Show();
            this.label7.Show();
            this.label8.Show();
            this.label9.Show();
            this.label10.Show();
            this.label11.Show();
            this.label13.Show();

            this.textBox1.Show();
            this.textBox2.Show();
            this.textBox3.Show();
            this.textBox4.Show();
            this.textBox5.Show();
            this.textBox6.Show();
            this.textBox7.Show();
            this.textBox8.Show();
            this.textBox9.Show();
            this.textBox10.Show();
            this.textBox11.Show();

            this.dataGridView1.Show();
            this.dataGridView2.Show();

            this.button1.Show();
            this.button2.Visible = false;
            this.button3.Visible = false;
            this.button4.Show();
            this.button5.Show();

           

            //set the tab order for all textboxes on this form

            this.textBox1.TabIndex = 0;
            this.textBox2.TabIndex = 1;
            this.textBox3.TabIndex = 2;
            this.textBox4.TabIndex = 3;
            this.textBox5.TabIndex = 4;
            this.textBox6.TabIndex = 5;
            this.textBox7.TabIndex = 6;
            this.textBox8.TabIndex = 7;
            this.textBox9.TabIndex = 8;
            this.textBox10.TabIndex = 9;
        }

        private void button4_Click(object sender, EventArgs e)//Create Copy
        {
            Form_AddImplementation form1 = new Form_AddImplementation();
            form1.Show();
            form1.Focus();

            //unhide the controls and put FileConsultantRequestForm.docx data on the controls of this form
            form1.Activate();
            form1.label12.Visible = false;
            form1.label1.Show();
            form1.label2.Show();
            form1.label3.Show();
            form1.label4.Show();
            form1.label5.Show();
            form1.label6.Show();
            form1.label7.Show();
            form1.label8.Show();
            form1.label9.Show();
            form1.label10.Show();
            form1.label11.Show();
            form1.label13.Show();

            form1.textBox1.Show();
            form1.textBox2.Show();
            form1.textBox3.Show();
            form1.textBox4.Show();
            form1.textBox5.Show();
            form1.textBox6.Show();
            form1.textBox7.Show();
            form1.textBox8.Show();
            form1.textBox9.Show();
            form1.textBox10.Show();
            form1.textBox11.Show();

            form1.dataGridView1.Show();
            form1.dataGridView2.Show();

            form1.button1.Show();
            form1.button2.Visible = false;
            form1.button3.Visible = false;
            form1.button4.Show();
            form1.button5.Show();



            //set the tab order for all textboxes on this form

            form1.textBox1.TabIndex = 0;
            form1.textBox2.TabIndex = 1;
            form1.textBox3.TabIndex = 2;
            form1.textBox4.TabIndex = 3;
            form1.textBox5.TabIndex = 4;
            form1.textBox6.TabIndex = 5;
            form1.textBox7.TabIndex = 6;
            form1.textBox8.TabIndex = 7;
            form1.textBox9.TabIndex = 8;
            form1.textBox10.TabIndex = 9;


            //import the data from one form to the other

            


            form1.textBox1.Text = this.textBox1.Text;
            form1.textBox2.Text = this.textBox2.Text;
            form1.textBox3.Text = this.textBox3.Text;
            form1.textBox4.Text = this.textBox4.Text;
            form1.textBox5.Text = this.textBox5.Text;
            form1.textBox6.Text = this.textBox6.Text;
            form1.textBox7.Text = this.textBox7.Text;
            form1.textBox8.Text = this.textBox8.Text;
            form1.textBox9.Text = this.textBox9.Text;
            form1.textBox10.Text = this.textBox10.Text;
            form1.textBox11.Text = this.textBox11.Text;





            //System.Data.DataTable dtForDataGridView1 = new System.Data.DataTable();

            //dtForDataGridView1.

            //this.dataGridView1

            //try
            //{
            //    for (int j = 0; j <= this.dataGridView1.ColumnCount; j++)
            //    {
            //        //form1.dataGridView1.Rows.Add();
            //        //this.dataGridView1.Rows.Add();
            //        for (int i = 0; i <= this.dataGridView1.Rows.Count; i++)
            //        {
            //            form1.dataGridView1[j, i].Value = this.dataGridView1[j, i].Value;
            //        }
            //    }
            //}
            //catch
            //{

            //}

            //try
            //{
            //    for (int j = 0; j <= this.dataGridView2.ColumnCount; j++)
            //    {
            //        //form1.dataGridView1.Rows.Add();
            //        //this.dataGridView1.Rows.Add();
            //        for (int i = 0; i <= this.dataGridView2.Rows.Count; i++)
            //        {
            //            form1.dataGridView2[j, i].Value = this.dataGridView2[j, i].Value;
            //        }
            //    }
            //}
            //catch
            //{

            //}


            try
            {
                form1.dataGridView1[0, 0].Value = this.dataGridView1[0, 0].Value;
                form1.dataGridView1[1, 0].Value = this.dataGridView1[1, 0].Value;
                form1.dataGridView1[2, 0].Value = this.dataGridView1[2, 0].Value;
                form1.dataGridView1[3, 0].Value = this.dataGridView1[3, 0].Value;

                //form1.dataGridView1[0, 1].Value = this.dataGridView1[0, 1].Value;
                //form1.dataGridView1[1, 1].Value = this.dataGridView1[1, 1].Value;
                //form1.dataGridView1[2, 1].Value = this.dataGridView1[2, 1].Value;
                //form1.dataGridView1[3, 1].Value = this.dataGridView1[3, 1].Value;

                //form1.dataGridView1[0, 2].Value = this.dataGridView1[0, 2].Value;
                //form1.dataGridView1[1, 2].Value = this.dataGridView1[1, 2].Value;
                //form1.dataGridView1[2, 2].Value = this.dataGridView1[2, 2].Value;
                //form1.dataGridView1[3, 2].Value = this.dataGridView1[3, 2].Value;             
            
                //form1.dataGridView1[0, 3].Value = this.dataGridView1[0, 3].Value;
                //form1.dataGridView1[1, 3].Value = this.dataGridView1[1, 3].Value;
                //form1.dataGridView1[2, 3].Value = this.dataGridView1[2, 3].Value;
                //form1.dataGridView1[3, 3].Value = this.dataGridView1[3, 3].Value;





                form1.dataGridView2[0, 0].Value = this.dataGridView2[0, 0].Value;
                form1.dataGridView2[1, 0].Value = this.dataGridView2[1, 0].Value;
                form1.dataGridView2[2, 0].Value = this.dataGridView2[2, 0].Value;
                form1.dataGridView2[3, 0].Value = this.dataGridView2[3, 0].Value;

                //form1.dataGridView2[0, 1].Value = this.dataGridView2[0, 1].Value;
                //form1.dataGridView2[1, 1].Value = this.dataGridView2[1, 1].Value;
                //form1.dataGridView2[2, 1].Value = this.dataGridView2[2, 1].Value;
                //form1.dataGridView2[3, 1].Value = this.dataGridView2[3, 1].Value;            
            
                //form1.dataGridView2[0, 2].Value = this.dataGridView2[0, 2].Value;
                //form1.dataGridView2[1, 2].Value = this.dataGridView2[1, 2].Value;
                //form1.dataGridView2[2, 2].Value = this.dataGridView2[2, 2].Value;
                //form1.dataGridView2[3, 2].Value = this.dataGridView2[3, 2].Value;
            
                //form1.dataGridView2[0, 3].Value = this.dataGridView2[0, 3].Value;
                //form1.dataGridView2[1, 3].Value = this.dataGridView2[1, 3].Value;
                //form1.dataGridView2[2, 3].Value = this.dataGridView2[2, 3].Value;
                //form1.dataGridView2[3, 3].Value = this.dataGridView2[3, 3].Value;




                ////datagridview1

                //if (dataGridView1.Rows.Count == 2)
                //{
                //    form1.dataGridView1[0, 1].Value = this.dataGridView1[0, 1].Value;
                //    form1.dataGridView1[1, 1].Value = this.dataGridView1[1, 1].Value;
                //    form1.dataGridView1[2, 1].Value = this.dataGridView1[2, 1].Value;
                //    form1.dataGridView1[3, 1].Value = this.dataGridView1[3, 1].Value;
                //}
                //else if (dataGridView1.Rows.Count == 3)
                //{
                //    form1.dataGridView1[0, 2].Value = this.dataGridView1[0, 2].Value;
                //    form1.dataGridView1[1, 2].Value = this.dataGridView1[1, 2].Value;
                //    form1.dataGridView1[2, 2].Value = this.dataGridView1[2, 2].Value;
                //    form1.dataGridView1[3, 2].Value = this.dataGridView1[3, 2].Value;
                //}
                //else if (dataGridView1.Rows.Count == 4)
                //{
                //    form1.dataGridView1[0, 3].Value = this.dataGridView1[0, 3].Value;
                //    form1.dataGridView1[1, 3].Value = this.dataGridView1[1, 3].Value;
                //    form1.dataGridView1[2, 3].Value = this.dataGridView1[2, 3].Value;
                //    form1.dataGridView1[3, 3].Value = this.dataGridView1[3, 3].Value;
                //}


                ////datagridview2
                //if (dataGridView2.Rows.Count == 2)
                //{

                //    form1.dataGridView2[0, 1].Value = this.dataGridView2[0, 1].Value;
                //    form1.dataGridView2[1, 1].Value = this.dataGridView2[1, 1].Value;
                //    form1.dataGridView2[2, 1].Value = this.dataGridView2[2, 1].Value;
                //    form1.dataGridView2[3, 1].Value = this.dataGridView2[3, 1].Value;
                //}
                //else if (dataGridView2.Rows.Count == 3)
                //{
                //    form1.dataGridView2[0, 2].Value = this.dataGridView2[0, 2].Value;
                //    form1.dataGridView2[1, 2].Value = this.dataGridView2[1, 2].Value;
                //    form1.dataGridView2[2, 2].Value = this.dataGridView2[2, 2].Value;
                //    form1.dataGridView2[3, 2].Value = this.dataGridView2[3, 2].Value;
                //}
                //else if (dataGridView2.Rows.Count == 4)
                //{
                //    form1.dataGridView2[0, 3].Value = this.dataGridView2[0, 3].Value;
                //    form1.dataGridView2[1, 3].Value = this.dataGridView2[1, 3].Value;
                //    form1.dataGridView2[2, 3].Value = this.dataGridView2[2, 3].Value;
                //    form1.dataGridView2[3, 3].Value = this.dataGridView2[3, 3].Value;
                //}
            }
            catch
            {

            }





        }

        private void button5_Click(object sender, EventArgs e)//create multiple copies
        {
            Form_HowManyFromCopiesPrompt formHowManyFromCopiesPrompt = new Form_HowManyFromCopiesPrompt();
            formHowManyFromCopiesPrompt.Show(); 


        }

        public void GetCopiesMethod()
        {
            int howManyCopies = Form_HowManyFromCopiesPrompt.HowMany;

            for (int i = 0; i < howManyCopies; i++)
            {
                Form_AddImplementation form1 = new Form_AddImplementation();
                form1.Show();
                form1.Focus();

                //unhide the controls and put FileConsultantRequestForm.docx data on the controls of this form
                form1.Activate();
                form1.label12.Visible = false;
                form1.label1.Show();
                form1.label2.Show();
                form1.label3.Show();
                form1.label4.Show();
                form1.label5.Show();
                form1.label6.Show();
                form1.label7.Show();
                form1.label8.Show();
                form1.label9.Show();
                form1.label10.Show();
                form1.label11.Show();
                form1.label13.Show();

                form1.textBox1.Show();
                form1.textBox2.Show();
                form1.textBox3.Show();
                form1.textBox4.Show();
                form1.textBox5.Show();
                form1.textBox6.Show();
                form1.textBox7.Show();
                form1.textBox8.Show();
                form1.textBox9.Show();
                form1.textBox10.Show();
                form1.textBox11.Show();

                form1.dataGridView1.Show();
                form1.dataGridView2.Show();

                form1.button1.Show();
                form1.button2.Visible = false;
                form1.button3.Visible = false;
                form1.button4.Show();
                form1.button5.Show();



                //set the tab order for all textboxes on this form

                form1.textBox1.TabIndex = 0;
                form1.textBox2.TabIndex = 1;
                form1.textBox3.TabIndex = 2;
                form1.textBox4.TabIndex = 3;
                form1.textBox5.TabIndex = 4;
                form1.textBox6.TabIndex = 5;
                form1.textBox7.TabIndex = 6;
                form1.textBox8.TabIndex = 7;
                form1.textBox9.TabIndex = 8;
                form1.textBox10.TabIndex = 9;


                //import the data from one form to the other




                form1.textBox1.Text = this.textBox1.Text;
                form1.textBox2.Text = this.textBox2.Text;
                form1.textBox3.Text = this.textBox3.Text;
                form1.textBox4.Text = this.textBox4.Text;
                form1.textBox5.Text = this.textBox5.Text;
                form1.textBox6.Text = this.textBox6.Text;
                form1.textBox7.Text = this.textBox7.Text;
                form1.textBox8.Text = this.textBox8.Text;
                form1.textBox9.Text = this.textBox9.Text;
                form1.textBox10.Text = this.textBox10.Text;
                form1.textBox11.Text = this.textBox11.Text;





               


                try
                {
                    form1.dataGridView1[0, 0].Value = this.dataGridView1[0, 0].Value;
                    form1.dataGridView1[1, 0].Value = this.dataGridView1[1, 0].Value;
                    form1.dataGridView1[2, 0].Value = this.dataGridView1[2, 0].Value;
                    form1.dataGridView1[3, 0].Value = this.dataGridView1[3, 0].Value;

                    //form1.dataGridView1[0, 1].Value = this.dataGridView1[0, 1].Value;
                    //form1.dataGridView1[1, 1].Value = this.dataGridView1[1, 1].Value;
                    //form1.dataGridView1[2, 1].Value = this.dataGridView1[2, 1].Value;
                    //form1.dataGridView1[3, 1].Value = this.dataGridView1[3, 1].Value;

                    //form1.dataGridView1[0, 2].Value = this.dataGridView1[0, 2].Value;
                    //form1.dataGridView1[1, 2].Value = this.dataGridView1[1, 2].Value;
                    //form1.dataGridView1[2, 2].Value = this.dataGridView1[2, 2].Value;
                    //form1.dataGridView1[3, 2].Value = this.dataGridView1[3, 2].Value;             

                    //form1.dataGridView1[0, 3].Value = this.dataGridView1[0, 3].Value;
                    //form1.dataGridView1[1, 3].Value = this.dataGridView1[1, 3].Value;
                    //form1.dataGridView1[2, 3].Value = this.dataGridView1[2, 3].Value;
                    //form1.dataGridView1[3, 3].Value = this.dataGridView1[3, 3].Value;





                    form1.dataGridView2[0, 0].Value = this.dataGridView2[0, 0].Value;
                    form1.dataGridView2[1, 0].Value = this.dataGridView2[1, 0].Value;
                    form1.dataGridView2[2, 0].Value = this.dataGridView2[2, 0].Value;
                    form1.dataGridView2[3, 0].Value = this.dataGridView2[3, 0].Value;

                    //form1.dataGridView2[0, 1].Value = this.dataGridView2[0, 1].Value;
                    //form1.dataGridView2[1, 1].Value = this.dataGridView2[1, 1].Value;
                    //form1.dataGridView2[2, 1].Value = this.dataGridView2[2, 1].Value;
                    //form1.dataGridView2[3, 1].Value = this.dataGridView2[3, 1].Value;            

                    //form1.dataGridView2[0, 2].Value = this.dataGridView2[0, 2].Value;
                    //form1.dataGridView2[1, 2].Value = this.dataGridView2[1, 2].Value;
                    //form1.dataGridView2[2, 2].Value = this.dataGridView2[2, 2].Value;
                    //form1.dataGridView2[3, 2].Value = this.dataGridView2[3, 2].Value;

                    //form1.dataGridView2[0, 3].Value = this.dataGridView2[0, 3].Value;
                    //form1.dataGridView2[1, 3].Value = this.dataGridView2[1, 3].Value;
                    //form1.dataGridView2[2, 3].Value = this.dataGridView2[2, 3].Value;
                    //form1.dataGridView2[3, 3].Value = this.dataGridView2[3, 3].Value;




                    ////datagridview1

                    //if (dataGridView1.Rows.Count == 2)
                    //{
                    //    form1.dataGridView1[0, 1].Value = this.dataGridView1[0, 1].Value;
                    //    form1.dataGridView1[1, 1].Value = this.dataGridView1[1, 1].Value;
                    //    form1.dataGridView1[2, 1].Value = this.dataGridView1[2, 1].Value;
                    //    form1.dataGridView1[3, 1].Value = this.dataGridView1[3, 1].Value;
                    //}
                    //else if (dataGridView1.Rows.Count == 3)
                    //{
                    //    form1.dataGridView1[0, 2].Value = this.dataGridView1[0, 2].Value;
                    //    form1.dataGridView1[1, 2].Value = this.dataGridView1[1, 2].Value;
                    //    form1.dataGridView1[2, 2].Value = this.dataGridView1[2, 2].Value;
                    //    form1.dataGridView1[3, 2].Value = this.dataGridView1[3, 2].Value;
                    //}
                    //else if (dataGridView1.Rows.Count == 4)
                    //{
                    //    form1.dataGridView1[0, 3].Value = this.dataGridView1[0, 3].Value;
                    //    form1.dataGridView1[1, 3].Value = this.dataGridView1[1, 3].Value;
                    //    form1.dataGridView1[2, 3].Value = this.dataGridView1[2, 3].Value;
                    //    form1.dataGridView1[3, 3].Value = this.dataGridView1[3, 3].Value;
                    //}


                    ////datagridview2
                    //if (dataGridView2.Rows.Count == 2)
                    //{

                    //    form1.dataGridView2[0, 1].Value = this.dataGridView2[0, 1].Value;
                    //    form1.dataGridView2[1, 1].Value = this.dataGridView2[1, 1].Value;
                    //    form1.dataGridView2[2, 1].Value = this.dataGridView2[2, 1].Value;
                    //    form1.dataGridView2[3, 1].Value = this.dataGridView2[3, 1].Value;
                    //}
                    //else if (dataGridView2.Rows.Count == 3)
                    //{
                    //    form1.dataGridView2[0, 2].Value = this.dataGridView2[0, 2].Value;
                    //    form1.dataGridView2[1, 2].Value = this.dataGridView2[1, 2].Value;
                    //    form1.dataGridView2[2, 2].Value = this.dataGridView2[2, 2].Value;
                    //    form1.dataGridView2[3, 2].Value = this.dataGridView2[3, 2].Value;
                    //}
                    //else if (dataGridView2.Rows.Count == 4)
                    //{
                    //    form1.dataGridView2[0, 3].Value = this.dataGridView2[0, 3].Value;
                    //    form1.dataGridView2[1, 3].Value = this.dataGridView2[1, 3].Value;
                    //    form1.dataGridView2[2, 3].Value = this.dataGridView2[2, 3].Value;
                    //    form1.dataGridView2[3, 3].Value = this.dataGridView2[3, 3].Value;
                    //}
                }
                catch
                {

                }
            }
        }


    }
}
