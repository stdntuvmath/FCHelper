using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Controls.Ribbon;
using System.Runtime.InteropServices;
using System.IO;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Outlook;
using System.Windows.Controls;
using System.Text.RegularExpressions;
using System.Windows.Input;


namespace FCHelper_v001
{



    public partial class Form_OpenImplementationList : Form
    {

        public static bool DuplicateForm = false;

        public static string EmployerName;
        public static string EmployerID;
        public static string Region;
        public static string Segment;
        public static string EffectiveDate;
        public static string CurrentProducts;
        public static string AddingProduct;
        public static string NewImp;
        public static string AMIMInvolved;
        public static string ImpDeadline;
        public static string SFTPCreds;

        public static string InternalContactName1;
        public static string InternalContactPhone1;
        public static string InternalContactEmail1;
        public static string InternalContactType1;

        public static string InternalContactName2;
        public static string InternalContactPhone2;
        public static string InternalContactEmail2;
        public static string InternalContactType2;

        public static string InternalContactName3;
        public static string InternalContactPhone3;
        public static string InternalContactEmail3;
        public static string InternalContactType3;

        public static string InternalContactName4;
        public static string InternalContactPhone4;
        public static string InternalContactEmail4;
        public static string InternalContactType4;


        public static string ExternalContactName1;
        public static string ExternalContactPhone1;
        public static string ExternalContactEmail1;
        public static string ExternalContactType1;

        public static string ExternalContactName2;
        public static string ExternalContactPhone2;
        public static string ExternalContactEmail2;
        public static string ExternalContactType2;

        public static string ExternalContactName3;
        public static string ExternalContactPhone3;
        public static string ExternalContactEmail3;
        public static string ExternalContactType3;

        public static string ExternalContactName4;
        public static string ExternalContactPhone4;
        public static string ExternalContactEmail4;
        public static string ExternalContactType4;


        //public static string SubjectLine;
        //public static string SubjectLine;
        //public static string SubjectLine;

        public static string SubjectLine;



        private string windowsUserName = System.Environment.UserName;//gives windows username

        

        // int currentRow;

        System.Data.DataTable dataTable1 = new System.Data.DataTable();//makes dataTable available
        System.Data.DataTable richTextBoxDatatable = new System.Data.DataTable();//makes dataTable available
        GetterSetterObject getObject = new GetterSetterObject();

        private bool button3WasClicked = false;
        private bool richtextboxTextWasChanged = false;


        public Form_OpenImplementationList()
        {
            InitializeComponent();


        }


        private void OpenImplementationList_Load(object sender, EventArgs e)
        {

            comboBox1.SelectedIndex = 0;

            dataTable1.Clear();
            this.Location = new System.Drawing.Point(60, 100);

            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();


            string excelDBFilePath = @"C:\Users\14025\Documents\File Consultants\ImpList.xls";

            FileInfo fso = new FileInfo(excelDBFilePath);

            //foreach (DataGridViewColumn column in dataGridView1.Columns)
            //{

            //    column.SortMode = DataGridViewColumnSortMode.Automatic;
            //}

            if (!fso.Exists)//if the file doesn't exist, creat a txt file and insert tab delimited headings into the text file
            {

                // Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

                Microsoft.Office.Interop.Excel.Workbook wb;
                Microsoft.Office.Interop.Excel.Worksheet ws;
                object misValue = System.Reflection.Missing.Value;

                wb = ExcelApp.Workbooks.Add(misValue);
                ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets.get_Item(1);


                //insert columns into spreadsheets
                ws.Name = "Brandon";
                ws.Cells[1, 1] = "ERname";
                ws.Cells[1, 2] = "ERID";
                ws.Cells[1, 3] = "Region";
                ws.Cells[1, 4] = "Segment";
                ws.Cells[1, 5] = "EffDate";
                ws.Cells[1, 6] = "CurrentProduct";
                ws.Cells[1, 7] = "AddingProduct";
                ws.Cells[1, 8] = "NewImplementation";
                ws.Cells[1, 9] = "AM_IM";
                ws.Cells[1, 10] = "ImplementationDeadline";
                ws.Cells[1, 11] = "SFTPFlag";
                ws.Cells[1, 12] = "InternalContactName";
                ws.Cells[1, 13] = "InternalContactPhone";
                ws.Cells[1, 14] = "InternalContactEmail";
                ws.Cells[1, 15] = "InternalContactType";
                ws.Cells[1, 16] = "ExternalContactName";
                ws.Cells[1, 17] = "ExternalContactPhone";
                ws.Cells[1, 18] = "ExternalContactEmail";
                ws.Cells[1, 19] = "ExternalContactType";
                ws.Cells[1, 20] = "FileType";
                ws.Cells[1, 21] = "chckbx1";
                ws.Cells[1, 22] = "chckbx2";
                ws.Cells[1, 23] = "chckbx3";
                ws.Cells[1, 24] = "chckbx4";
                ws.Cells[1, 25] = "chckbx5";
                ws.Cells[1, 26] = "chckbx6";
                ws.Cells[1, 27] = "chckbx7";
                ws.Cells[1, 28] = "chckbx8";
                ws.Cells[1, 29] = "chckbx9";
                ws.Cells[1, 30] = "chckbx10";
                ws.Cells[1, 31] = "notes";


                wb.SaveAs(excelDBFilePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlShared, misValue, misValue, misValue, misValue, misValue);
                wb.Close(true, misValue, misValue);
                ExcelApp.Quit();

            }

            string excelTableSource = "Provider = Microsoft.Jet.OLEDB.4.0; Data Source = " + excelDBFilePath + "; Extended Properties = 'Excel 8.0; HDR = YES'";

            OleDbConnection connection = new OleDbConnection();
            OleDbCommand command = new OleDbCommand();

            connection = new OleDbConnection(excelTableSource);

            string sql = "SELECT * FROM [Brandon$]";
            OleDbDataAdapter fillFromBrandon = new OleDbDataAdapter(sql, connection);//pulls data from Brandon in excel doc to use Fill method                         

            connection.Open();
            command.Connection = connection;

            command.CommandText = sql;
            command.ExecuteNonQuery();

            fillFromBrandon.Fill(dataTable1);

            connection.Close();
            sql = null;

            ExcelApp.Quit();
            //ExcelApp = null;

            dataGridView1.DataSource = dataTable1;

            //load the initially selected row into the textBoxs


            if (dataGridView1.Rows.Count > 0 && dataGridView1.SelectedRows != null)
            {
                int currentRow = dataGridView1.CurrentRow.Index;
                this.textBox1.Text = dataGridView1[0, currentRow].Value.ToString();
                this.textBox2.Text = dataGridView1[1, currentRow].Value.ToString();
                this.textBox3.Text = dataGridView1[2, currentRow].Value.ToString();
                this.textBox4.Text = dataGridView1[3, currentRow].Value.ToString();
                this.textBox5.Text = dataGridView1[4, currentRow].Value.ToString();
                this.textBox6.Text = dataGridView1[5, currentRow].Value.ToString();
                this.textBox7.Text = dataGridView1[6, currentRow].Value.ToString();
                this.textBox8.Text = dataGridView1[7, currentRow].Value.ToString();
                this.textBox9.Text = dataGridView1[8, currentRow].Value.ToString();
                this.textBox10.Text = dataGridView1[9, currentRow].Value.ToString();

                EmployerID = textBox2.Text;


                //add internal contacts to contact boxes


                dataGridView2.Rows.Add();
                dataGridView2.Rows.Add();
                dataGridView2.Rows.Add();

                dataGridView3.Rows.Add();
                dataGridView3.Rows.Add();
                dataGridView3.Rows.Add();


                dataGridView2[0, 0].Value = dataGridView1[11, currentRow].Value.ToString();
                dataGridView2[1,0].Value = dataGridView1[13, currentRow].Value.ToString();
                dataGridView2[2,0].Value = dataGridView1[12, currentRow].Value.ToString();
                dataGridView2[3, 0].Value = dataGridView1[14, currentRow].Value.ToString();

                dataGridView2[0, 1].Value = dataGridView1[31, currentRow].Value.ToString();
                dataGridView2[1, 1].Value = dataGridView1[33, currentRow].Value.ToString();
                dataGridView2[2, 1].Value = dataGridView1[32, currentRow].Value.ToString();
                dataGridView2[3, 1].Value = dataGridView1[34, currentRow].Value.ToString();

                dataGridView2[0, 2].Value = dataGridView1[35, currentRow].Value.ToString();
                dataGridView2[1, 2].Value = dataGridView1[37, currentRow].Value.ToString();
                dataGridView2[2, 2].Value = dataGridView1[36, currentRow].Value.ToString();
                dataGridView2[3, 2].Value = dataGridView1[38, currentRow].Value.ToString();

                dataGridView2[0, 3].Value = dataGridView1[39, currentRow].Value.ToString();
                dataGridView2[1, 3].Value = dataGridView1[41, currentRow].Value.ToString();
                dataGridView2[2, 3].Value = dataGridView1[40, currentRow].Value.ToString();
                dataGridView2[3, 3].Value = dataGridView1[42, currentRow].Value.ToString();



                //add external contacts to contact boxes
                dataGridView3[0, 0].Value = dataGridView1[15, currentRow].Value.ToString();
                dataGridView3[1, 0].Value = dataGridView1[17, currentRow].Value.ToString();
                dataGridView3[2, 0].Value = dataGridView1[16, currentRow].Value.ToString();
                dataGridView3[3, 0].Value = dataGridView1[18, currentRow].Value.ToString();

                dataGridView3[0, 1].Value = dataGridView1[43, currentRow].Value.ToString();
                dataGridView3[1, 1].Value = dataGridView1[45, currentRow].Value.ToString();
                dataGridView3[2, 1].Value = dataGridView1[44, currentRow].Value.ToString();
                dataGridView3[3, 1].Value = dataGridView1[46, currentRow].Value.ToString();

                dataGridView3[0, 2].Value = dataGridView1[47, currentRow].Value.ToString();
                dataGridView3[1, 2].Value = dataGridView1[49, currentRow].Value.ToString();
                dataGridView3[2, 2].Value = dataGridView1[48, currentRow].Value.ToString();
                dataGridView3[3, 2].Value = dataGridView1[50, currentRow].Value.ToString();

                dataGridView3[0, 3].Value = dataGridView1[51, currentRow].Value.ToString();
                dataGridView3[1, 3].Value = dataGridView1[53, currentRow].Value.ToString();
                dataGridView3[2, 3].Value = dataGridView1[52, currentRow].Value.ToString();
                dataGridView3[3, 3].Value = dataGridView1[54, currentRow].Value.ToString();

            }
            else
            {
                this.textBox1.Text = "";
                this.textBox2.Text = "";
                this.textBox3.Text = "";
                this.textBox4.Text = "";
                this.textBox5.Text = "";
                this.textBox6.Text = "";
                this.textBox7.Text = "";
                this.textBox8.Text = "";
                this.textBox9.Text = "";
                this.textBox10.Text = "";


            }
            //set first tab index
            dataGridView1.TabIndex = 1;

            //set column and row start
            this.dataGridView1.CurrentCell.Selected = false;
            this.dataGridView1.CurrentCell = this.dataGridView1.Rows[0].Cells[0];



            //kill excel process generated by this form load
            ExcelApp.Quit();

            dataGridView1.Rows[dataGridView1.CurrentRow.Index].Selected = true;
        }

        private void OpenImplementationList_FormClosed(object sender, FormClosedEventArgs e)
        {
            try
            {
                foreach (Process proc in Process.GetProcessesByName("excel"))
                {
                    
                    proc.Kill();
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void DataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {

            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();

            // Update the DB with new data.
            int currentRow = dataGridView1.CurrentCell.RowIndex;
            int currentColumn = dataGridView1.CurrentCell.ColumnIndex;
            string column = "";

            string locationColumn = "ERID";
            string erid = textBox2.Text;





            if (currentColumn == 0)
            {
                

                column = "ERname";
                textBox1.Text = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);

               

            }
            else if (currentColumn == 1)
            {
                column = "ERID";
                textBox2.Text = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);

            }
            else if (currentColumn == 2)
            {
                column = "Region";
                textBox3.Text = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);

            }
            else if (currentColumn == 3)
            {
                column = "Segment";
                textBox4.Text = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);

            }
            else if (currentColumn == 4)
            {
                column = "EffDate";
                textBox5.Text = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);

            }
            else if (currentColumn == 5)
            {
                column = "CurrentProduct";
                textBox6.Text = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);

            }
            else if (currentColumn == 6)
            {
                column = "AddingProduct";
                textBox7.Text = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);

            }
            else if (currentColumn == 7)
            {
                column = "NewImplementation";
                textBox8.Text = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);

            }
            else if (currentColumn == 8)
            {
                column = "AM_IM";
                textBox9.Text = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);

            }
            else if (currentColumn == 9)
            {
                column = "ImplementationDeadline";
                textBox10.Text = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);

            }
            else if (currentColumn == 10)
            {
                column = "SFTPFlag";
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);


            }

            else if (currentColumn == 11)
            {
                column = "InternalContactName";
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);


            }
            else if (currentColumn == 12)
            {
                column = "InternalContactPhone";
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);


            }
            else if (currentColumn == 13)
            {
                column = "InternalContactEmail";
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);


            }
            else if (currentColumn == 14)
            {
                column = "InternalContactType";
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);


            }
            else if (currentColumn == 15)
            {
                column = "ExternalContactName";
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);


            }
            else if (currentColumn == 16)
            {
                column = "ExternalContactPhone";
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);


            }
            else if (currentColumn == 17)
            {
                column = "ExternalContactEmail";
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);


            }
            else if (currentColumn == 18)
            {
                column = "ExternalContactType";
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);


            }
            else if (currentColumn == 19)
            {
                column = "FileType";
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);


            }
            else if (currentColumn == 20)
            {
                column = "chckbx1";
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);


            }
            else if (currentColumn == 21)
            {
                column = "chckbx2";
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);


            }
            else if (currentColumn == 22)
            {
                column = "chckbx3";
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);


            }
            else if (currentColumn == 23)
            {
                column = "chckbx4";
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);


            }
            else if (currentColumn == 24)
            {
                column = "chckbx5";
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);


            }
            else if (currentColumn == 25)
            {
                column = "chckbx6";
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);


            }
            else if (currentColumn == 26)
            {
                column = "chckbx7";
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);


            }
            else if (currentColumn == 27)
            {
                column = "chckbx8";
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);


            }
            else if (currentColumn == 28)
            {
                column = "chckbx9";
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);


            }
            else if (currentColumn == 29)
            {
                column = "chckbx10";
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);


            }
            else if (currentColumn == 30)
            {
                column = "notes";
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);


            }
            else if (currentColumn == 31)
            {
                column = "InternalContactName2";
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);


            }
            else if (currentColumn == 32)
            {
                column = "InternalContactPhone2";
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);


            }
            else if (currentColumn == 33)
            {
                column = "InternalContactEmail2";
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);


            }
            else if (currentColumn == 34)
            {
                column = "InternalContactType2";
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);


            }
            else if (currentColumn == 35)
            {
                column = "InternalContactName3";
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);


            }
            else if (currentColumn == 36)
            {
                column = "InternalContactPhone3";
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);


            }
            else if (currentColumn == 37)
            {
                column = "InternalContactEmail3";
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);


            }
            else if (currentColumn == 38)
            {
                column = "InternalContactType3";
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);


            }
            else if (currentColumn == 39)
            {
                column = "InternalContactName4";
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);


            }
            else if (currentColumn == 40)
            {
                column = "InternalContactPhone4";
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);


            }
            else if (currentColumn == 41)
            {
                column = "InternalContactEmail4";
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);


            }
            else if (currentColumn == 42)
            {
                column = "InternalContactType4";
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);


            }
            else if (currentColumn == 43)
            {
                column = "ExternalContactName2";
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);


            }
            else if (currentColumn == 44)
            {
                column = "ExternalContactPhone2";
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);


            }
            else if (currentColumn == 45)
            {
                column = "ExternalContactEmail2";
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);


            }
            else if (currentColumn == 46)
            {
                column = "ExternalContactType2";
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);


            }
            else if (currentColumn == 47)
            {
                column = "ExternalContactName3";
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);


            }
            else if (currentColumn == 48)
            {
                column = "ExternalContactPhone3";
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);


            }
            else if (currentColumn == 49)
            {
                column = "ExternalContactEmail3";
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);


            }
            else if (currentColumn == 50)
            {
                column = "ExternalContactType3";
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);


            }
            else if (currentColumn == 51)
            {
                column = "ExternalContactName4";
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);


            }
            else if (currentColumn == 52)
            {
                column = "ExternalContactPhone4";
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);


            }
            else if (currentColumn == 53)
            {
                column = "ExternalContactEmail4";
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);


            }
            else if (currentColumn == 54)
            {
                column = "ExternalContactType4";
                string value = "";
                value = dataGridView1.Rows[currentRow].Cells[currentColumn].Value.ToString();

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);


            }

            ExcelApp.Quit();
            dataGridView1.Rows[dataGridView1.CurrentRow.Index].Selected = true;
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {

            listView1.Clear();
            dataGridView1.Select();
            //int index = dataGridView1.CurrentRow.Index;

            //if (dataGridView1.SelectedRows[index].ToString() == "")
            //{
            //    dataGridView1.Rows.RemoveAt(index);
            //}
            //take care of null column click issue (null index)
            //if (dataGridView1.currentRow != null && dataGridView1.currentRow.Index < dataGridView1.Rows.Count - 1)
            //{
            //    int currentRow = dataGridView1.currentRow.Index + 1;
            //}
            //else
            //{


            //}

            //reset richtextbox colors to default colors
            richTextBox1.BackColor = Color.White;
            richTextBox1.ForeColor = Color.Black;

            int currentRow = dataGridView1.CurrentRow.Index;


            //richTextBox1.Text = dataGridView1[30, currentRow].Value.ToString();


            //label15.Text = "Total Count: "+dataGridView1.RowCount.ToString();





            //clear all previous checks in checklistbox
            foreach (int i in checkedListBox1.CheckedIndices)
            {
                checkedListBox1.SetItemCheckState(i, CheckState.Unchecked);
            }





            // int currentRow = dataGridView1.CurrentCell.RowIndex;
            //check the checkboxes from the datagridview data
            if (dataGridView1[20, currentRow].Value.ToString() == "1")
            {
                checkedListBox1.SetItemCheckState(0, CheckState.Checked);
            }
            else if (dataGridView1[20, currentRow].Value.ToString() == "0")
            {
                checkedListBox1.SetItemCheckState(0, CheckState.Unchecked);
            }


            if (dataGridView1[21, currentRow].Value.ToString() == "1")
            {
                checkedListBox1.SetItemCheckState(1, CheckState.Checked);
            }
            else if (dataGridView1[21, currentRow].Value.ToString() == "0")
            {
                checkedListBox1.SetItemCheckState(1, CheckState.Unchecked);
            }


            if (dataGridView1[22, currentRow].Value.ToString() == "1")
            {
                checkedListBox1.SetItemCheckState(2, CheckState.Checked);
            }
            else if (dataGridView1[22, currentRow].Value.ToString() == "0")
            {
                checkedListBox1.SetItemCheckState(2, CheckState.Unchecked);
            }

            if (dataGridView1[23, currentRow].Value.ToString() == "1")
            {
                checkedListBox1.SetItemCheckState(3, CheckState.Checked);
            }
            else if (dataGridView1[23, currentRow].Value.ToString() == "0")
            {
                checkedListBox1.SetItemCheckState(3, CheckState.Unchecked);
            }


            if (dataGridView1[24, currentRow].Value.ToString() == "1")
            {
                checkedListBox1.SetItemCheckState(4, CheckState.Checked);
            }
            else if (dataGridView1[24, currentRow].Value.ToString() == "0")
            {
                checkedListBox1.SetItemCheckState(4, CheckState.Unchecked);
            }


            if (dataGridView1[25, currentRow].Value.ToString() == "1")
            {
                checkedListBox1.SetItemCheckState(5, CheckState.Checked);
            }
            else if (dataGridView1[25, currentRow].Value.ToString() == "0")
            {
                checkedListBox1.SetItemCheckState(5, CheckState.Unchecked);
            }


            if (dataGridView1[26, currentRow].Value.ToString() == "1")
            {
                checkedListBox1.SetItemCheckState(6, CheckState.Checked);
            }
            else if (dataGridView1[26, currentRow].Value.ToString() == "0")
            {
                checkedListBox1.SetItemCheckState(6, CheckState.Unchecked);
            }


            if (dataGridView1[27, currentRow].Value.ToString() == "1")
            {
                checkedListBox1.SetItemCheckState(7, CheckState.Checked);
            }
            else if (dataGridView1[27, currentRow].Value.ToString() == "0")
            {
                checkedListBox1.SetItemCheckState(7, CheckState.Unchecked);
            }


            if (dataGridView1[28, currentRow].Value.ToString() == "1")
            {
                checkedListBox1.SetItemCheckState(8, CheckState.Checked);
            }
            else if (dataGridView1[28, currentRow].Value.ToString() == "0")
            {
                checkedListBox1.SetItemCheckState(8, CheckState.Unchecked);
            }


            if (dataGridView1[29, currentRow].Value.ToString() == "1")
            {
                checkedListBox1.SetItemCheckState(9, CheckState.Checked);
            }
            else if (dataGridView1[29, currentRow].Value.ToString() == "0")
            {
                checkedListBox1.SetItemCheckState(9, CheckState.Unchecked);
            }





            //set the progress bar with the amount of checked boxes
            progressBar1.Maximum = 100;
            progressBar1.Style = System.Windows.Forms.ProgressBarStyle.Blocks;

            switch (checkedListBox1.CheckedItems.Count)
            {
                case 0:
                    progressBar1.Value = 0;
                    label12.Text = "Status: 0%";
                    break;
                case 1:
                    progressBar1.Value = 10;
                    progressBar1.ForeColor = Color.Red;
                    label12.Text = "Status: 10%";
                    break;
                case 2:
                    progressBar1.Value = 20;
                    progressBar1.ForeColor = Color.OrangeRed;
                    label12.Text = "Status: 20%";
                    break;
                case 3:
                    progressBar1.Value = 30;
                    progressBar1.ForeColor = Color.DarkOrange;
                    label12.Text = "Status: 30%";
                    break;
                case 4:
                    progressBar1.Value = 40;
                    progressBar1.ForeColor = Color.Orange;
                    label12.Text = "Status: 40%";
                    break;
                case 5:
                    progressBar1.Value = 50;
                    progressBar1.ForeColor = Color.Gold;
                    label12.Text = "Status: 50%";
                    break;
                case 6:
                    System.Drawing.Color someYellowColor = System.Drawing.ColorTranslator.FromHtml("#f4c741");
                    progressBar1.Value = 60;
                    progressBar1.ForeColor = someYellowColor;
                    label12.Text = "Status: 60%";
                    break;
                case 7:
                    System.Drawing.Color someGreenColor = System.Drawing.ColorTranslator.FromHtml("#ebf441");

                    progressBar1.Value = 70;
                    progressBar1.ForeColor = someGreenColor;
                    label12.Text = "Status: 70%";
                    break;
                case 8:
                    progressBar1.Value = 80;
                    progressBar1.ForeColor = Color.GreenYellow;
                    label12.Text = "Status: 80%";
                    break;
                case 9:
                    progressBar1.Value = 90;
                    progressBar1.ForeColor = Color.LawnGreen;
                    label12.Text = "Status: 90%";
                    break;
                case 10:
                    progressBar1.Value = 100;
                    progressBar1.ForeColor = Color.Lime;
                    label12.Text = "Status: 100%";
                    break;

            }




            //load the initially selected row into the textBoxs

            if (dataGridView1[0, currentRow].Value.ToString() != string.Empty)
            {
                this.textBox1.Text = (string)dataGridView1[0, currentRow].Value;

            }
            else if (dataGridView1[0, currentRow].Value.ToString() == string.Empty)
            {
                this.textBox1.Text = "";
            }

            if (dataGridView1[1, currentRow].Value.ToString() != string.Empty)
            {
                this.textBox2.Text = (string)dataGridView1[1, currentRow].Value;

            }
            else if (dataGridView1[1, currentRow].Value.ToString() == string.Empty)
            {
                this.textBox2.Text = "";
            }

            if (dataGridView1[2, currentRow].Value.ToString() != string.Empty)
            {
                this.textBox3.Text = (string)dataGridView1[2, currentRow].Value;

            }
            else if (dataGridView1[2, currentRow].Value.ToString() == string.Empty)
            {
                this.textBox3.Text = "";
            }

            if (dataGridView1[3, currentRow].Value.ToString() != string.Empty)
            {
                this.textBox4.Text = (string)dataGridView1[3, currentRow].Value;

            }
            else if (dataGridView1[3, currentRow].Value.ToString() == string.Empty)
            {
                this.textBox4.Text = "";
            }

            if (dataGridView1[4, currentRow].Value.ToString() != string.Empty)
            {
                this.textBox5.Text = (string)dataGridView1[4, currentRow].Value;

            }
            else if (dataGridView1[4, currentRow].Value.ToString() == string.Empty)
            {
                this.textBox5.Text = "";
            }

            if (dataGridView1[5, currentRow].Value.ToString() != string.Empty)
            {
                this.textBox6.Text = (string)dataGridView1[5, currentRow].Value;

            }
            else if (dataGridView1[5, currentRow].Value.ToString() == string.Empty)
            {
                this.textBox6.Text = "";
            }

            if (dataGridView1[6, currentRow].Value.ToString() != string.Empty)
            {
                this.textBox7.Text = (string)dataGridView1[6, currentRow].Value;

            }
            else if (dataGridView1[6, currentRow].Value.ToString() == string.Empty)
            {
                this.textBox7.Text = "";
            }

            if (dataGridView1[7, currentRow].Value.ToString() != string.Empty)
            {
                this.textBox8.Text = (string)dataGridView1[7, currentRow].Value;

            }
            else if (dataGridView1[7, currentRow].Value.ToString() == string.Empty)
            {
                this.textBox8.Text = "";
            }

            if (dataGridView1[8, currentRow].Value.ToString() != string.Empty)
            {
                this.textBox9.Text = (string)dataGridView1[8, currentRow].Value;

            }
            else if (dataGridView1[8, currentRow].Value.ToString() == string.Empty)
            {
                this.textBox9.Text = "";
            }

            if (dataGridView1[9, currentRow].Value.ToString() != string.Empty)
            {
                this.textBox10.Text = (string)dataGridView1[9, currentRow].Value;

            }
            else if (dataGridView1[9, currentRow].Value.ToString() == string.Empty)
            {
                this.textBox10.Text = "";
            }

            if (dataGridView1[10, currentRow].Value.ToString() != string.Empty)
            {
                this.textBox11.Text = (string)dataGridView1[10, currentRow].Value;

            }
            else if (dataGridView1[10, currentRow].Value.ToString() == string.Empty)
            {
                this.textBox11.Text = "";
            }

            if (dataGridView1[11, currentRow].Value.ToString() != string.Empty)
            {
                this.textBox11.Text = (string)dataGridView1[10, currentRow].Value;

            }
            else if (dataGridView1[11, currentRow].Value.ToString() == string.Empty)
            {
                this.textBox11.Text = "";
            }

            dataGridView2.Rows.Clear();
            dataGridView3.Rows.Clear();

            dataGridView2.Rows.Add();
            dataGridView2.Rows.Add();
            dataGridView2.Rows.Add();
            dataGridView2.Rows.Add();

            dataGridView3.Rows.Add();
            dataGridView3.Rows.Add();
            dataGridView3.Rows.Add();
            dataGridView3.Rows.Add();


            dataGridView2[0, 0].Value = dataGridView1[11, currentRow].Value.ToString();
            dataGridView2[1, 0].Value = dataGridView1[13, currentRow].Value.ToString();
            dataGridView2[2, 0].Value = dataGridView1[12, currentRow].Value.ToString();
            dataGridView2[3, 0].Value = dataGridView1[14, currentRow].Value.ToString();

            dataGridView2[0, 1].Value = dataGridView1[31, currentRow].Value.ToString();
            dataGridView2[1, 1].Value = dataGridView1[33, currentRow].Value.ToString();
            dataGridView2[2, 1].Value = dataGridView1[32, currentRow].Value.ToString();
            dataGridView2[3, 1].Value = dataGridView1[34, currentRow].Value.ToString();

            dataGridView2[0, 2].Value = dataGridView1[35, currentRow].Value.ToString();
            dataGridView2[1, 2].Value = dataGridView1[37, currentRow].Value.ToString();
            dataGridView2[2, 2].Value = dataGridView1[36, currentRow].Value.ToString();
            dataGridView2[3, 2].Value = dataGridView1[38, currentRow].Value.ToString();

            dataGridView2[0, 3].Value = dataGridView1[39, currentRow].Value.ToString();
            dataGridView2[1, 3].Value = dataGridView1[41, currentRow].Value.ToString();
            dataGridView2[2, 3].Value = dataGridView1[40, currentRow].Value.ToString();
            dataGridView2[3, 3].Value = dataGridView1[42, currentRow].Value.ToString();



            //add external contacts to contact boxes
            dataGridView3[0, 0].Value = dataGridView1[15, currentRow].Value.ToString();
            dataGridView3[1, 0].Value = dataGridView1[17, currentRow].Value.ToString();
            dataGridView3[2, 0].Value = dataGridView1[16, currentRow].Value.ToString();
            dataGridView3[3, 0].Value = dataGridView1[18, currentRow].Value.ToString();

            dataGridView3[0, 1].Value = dataGridView1[43, currentRow].Value.ToString();
            dataGridView3[1, 1].Value = dataGridView1[45, currentRow].Value.ToString();
            dataGridView3[2, 1].Value = dataGridView1[44, currentRow].Value.ToString();
            dataGridView3[3, 1].Value = dataGridView1[46, currentRow].Value.ToString();

            dataGridView3[0, 2].Value = dataGridView1[47, currentRow].Value.ToString();
            dataGridView3[1, 2].Value = dataGridView1[49, currentRow].Value.ToString();
            dataGridView3[2, 2].Value = dataGridView1[48, currentRow].Value.ToString();
            dataGridView3[3, 2].Value = dataGridView1[50, currentRow].Value.ToString();

            dataGridView3[0, 3].Value = dataGridView1[51, currentRow].Value.ToString();
            dataGridView3[1, 3].Value = dataGridView1[53, currentRow].Value.ToString();
            dataGridView3[2, 3].Value = dataGridView1[52, currentRow].Value.ToString();
            dataGridView3[3, 3].Value = dataGridView1[54, currentRow].Value.ToString();

            //ImageList1.Images.Add(Icon.ExtractAssociatedIcon(fileName))
            //System.Drawing.Icon.ExtractAssociatedIcon(completeFilePath);




            //filter textbox1 and 2 to remove carriage returns

            textBox1.Text = textBox1.Text.TrimEnd('\r', '\n');
            textBox2.Text = textBox2.Text.TrimEnd('\r', '\n');

            //load richtextbox data from employer specific text file

            string file = @"C:\Users\14025\Documents\File Consultants\Brandon\Notes\" + textBox1.Text + "_" + textBox2.Text + "_Notes.rtf";
            if (File.Exists(file))
            {
                try
                {
                    using (var rtf = new RichTextBox())
                    {
                        rtf.Rtf = File.ReadAllText(file);
                        //return rtf.Text;
                        richTextBox1.Rtf = rtf.Rtf;
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show("Method: dataGridView1_SelectionChanged\rSomething prevented the file from loading to the RickTextBox.\r\r" + ex, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
            }
            else
            {
                DialogResult result = MessageBox.Show("Notes file does not exist. Would you like to create one?", "Create Notes?", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);


                if (result == DialogResult.Yes)
                {

                    richTextBox1.Clear();
                    richTextBox1.SelectionFont = new System.Drawing.Font(richTextBox1.Font, FontStyle.Bold);
                    richTextBox1.AppendText("Current Status:");

                    richTextBox1.SelectionFont = new System.Drawing.Font(richTextBox1.Font, FontStyle.Regular);
                    richTextBox1.AppendText(" ");

                    string textToSave = richTextBox1.Rtf;

                    //get a safe filename



                    //MessageBox.Show(file + " has been successfully saved to the Notes folder.");

                    SaveAndWriteToRTFFile save = new SaveAndWriteToRTFFile();
                    save.SaveAndWriteToRTFFileMethod(textToSave, file);
                }
                else if (result == DialogResult.No)
                {
                    //do nothing
                }
            }
            //get fies in empoyer implementaion folder


            string filePath = @"C:\Users\14025\Documents\File Consultants\Brandon\";
            string erNameERID = textBox1.Text + "_" + textBox2.Text + @"\";
            string completeFilePath = filePath + erNameERID;

            if (!Directory.Exists(completeFilePath))
            {
                try
                {
                    Directory.CreateDirectory(completeFilePath);
                }
                catch (System.IO.IOException ex)
                {
                    MessageBox.Show(ex.ToString());
                }
                catch (System.ArgumentException ex)
                {
                    MessageBox.Show(ex.ToString());
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }

            }

            //MessageBox.Show(completeFilePath);

            try
            {
                var filesInDirectory = Directory.EnumerateFiles(completeFilePath);

                //MessageBox.Show(filesInDirectory.ToString());

                
                

                if (filesInDirectory != null)
                {
                    string[] files = Directory.GetFiles(completeFilePath);
                    int counter = 0;
                    foreach (string file1 in files)
                    {

                        counter++;
                        string filePath1 = System.IO.Path.GetDirectoryName(file1);
                        string fileName = System.IO.Path.GetFileName(file1);

                        listView1.Items.Add(fileName);
                    }
                    //set running total labels
                    label16.Text = "Files: " + counter;
                    label15.Text = "Implementations: " + dataGridView1.Rows.Count;
                    EmployerID = null;
                    EmployerID = textBox2.Text;
                }
                else
                {

                }
               

            }
            catch(System.IO.IOException ex)
            {
                MessageBox.Show("Error in method dataGridView1_SelectionChanged()\r\r"+ex, "Could not GetFiles()",MessageBoxButtons.OK,MessageBoxIcon.Error);
            }
            catch (System.ArgumentException ex)
            {
                MessageBox.Show("Error in method dataGridView1_SelectionChanged()\r\r" + ex, "Could not GetFiles()", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }


            //dataGridView1.Rows[dataGridView1.CurrentRow.Index].Selected = true;

        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();


            //backgroundWorker1.RunWorkerAsync();

            //save checkedListBox1 data to datagridview1

            int currentRow = dataGridView1.CurrentCell.RowIndex;

            if (checkedListBox1.GetItemCheckState(0) == CheckState.Checked)
            {
                dataGridView1[20, currentRow].Value = "1";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx1";
                string value = "1";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }
            else if (checkedListBox1.GetItemCheckState(0) == CheckState.Unchecked)
            {
                dataGridView1[20, currentRow].Value = "0";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx1";
                string value = "0";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }

            if (checkedListBox1.GetItemCheckState(1) == CheckState.Checked)
            {
                dataGridView1[21, currentRow].Value = "1";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx2";
                string value = "1";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }
            else if (checkedListBox1.GetItemCheckState(1) == CheckState.Unchecked)
            {
                dataGridView1[21, currentRow].Value = "0";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx2";
                string value = "0";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }
            if (checkedListBox1.GetItemCheckState(2) == CheckState.Checked)
            {
                dataGridView1[22, currentRow].Value = "1";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx3";
                string value = "1";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }
            else if (checkedListBox1.GetItemCheckState(2) == CheckState.Unchecked)
            {
                dataGridView1[22, currentRow].Value = "0";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx3";
                string value = "0";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }

            if (checkedListBox1.GetItemCheckState(3) == CheckState.Checked)
            {
                dataGridView1[23, currentRow].Value = "1";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx4";
                string value = "1";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }
            else if (checkedListBox1.GetItemCheckState(3) == CheckState.Unchecked)
            {
                dataGridView1[23, currentRow].Value = "0";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx4";
                string value = "0";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }

            if (checkedListBox1.GetItemCheckState(4) == CheckState.Checked)
            {
                dataGridView1[24, currentRow].Value = "1";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx5";
                string value = "1";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }
            else if (checkedListBox1.GetItemCheckState(4) == CheckState.Unchecked)
            {
                dataGridView1[24, currentRow].Value = "0";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx5";
                string value = "0";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }

            if (checkedListBox1.GetItemCheckState(5) == CheckState.Checked)
            {
                dataGridView1[25, currentRow].Value = "1";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx6";
                string value = "1";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }
            else if (checkedListBox1.GetItemCheckState(5) == CheckState.Unchecked)
            {
                dataGridView1[25, currentRow].Value = "0";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx6";
                string value = "0";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }

            if (checkedListBox1.GetItemCheckState(6) == CheckState.Checked)
            {
                dataGridView1[26, currentRow].Value = "1";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx7";
                string value = "1";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }
            else if (checkedListBox1.GetItemCheckState(6) == CheckState.Unchecked)
            {
                dataGridView1[26, currentRow].Value = "0";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx7";
                string value = "0";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }

            if (checkedListBox1.GetItemCheckState(7) == CheckState.Checked)
            {
                dataGridView1[27, currentRow].Value = "1";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx8";
                string value = "1";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }
            else if (checkedListBox1.GetItemCheckState(7) == CheckState.Unchecked)
            {
                dataGridView1[27, currentRow].Value = "0";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx8";
                string value = "0";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }

            if (checkedListBox1.GetItemCheckState(8) == CheckState.Checked)
            {
                dataGridView1[28, currentRow].Value = "1";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx9";
                string value = "1";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }
            else if (checkedListBox1.GetItemCheckState(8) == CheckState.Unchecked)
            {
                dataGridView1[28, currentRow].Value = "0";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx9";
                string value = "0";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }

            if (checkedListBox1.GetItemCheckState(9) == CheckState.Checked)
            {
                dataGridView1[29, currentRow].Value = "1";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx10";
                string value = "1";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }
            else if (checkedListBox1.GetItemCheckState(9) == CheckState.Unchecked)
            {
                dataGridView1[29, currentRow].Value = "0";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx10";
                string value = "0";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }


            //set the progress bar with the amount of checked boxes
            progressBar1.Maximum = 100;
            progressBar1.Style = System.Windows.Forms.ProgressBarStyle.Blocks;

            switch (checkedListBox1.CheckedItems.Count)
            {
                case 0:
                    progressBar1.Value = 0;
                    label12.Text = "Status: 0%";
                    break;
                case 1:
                    progressBar1.Value = 10;
                    progressBar1.ForeColor = Color.Red;
                    label12.Text = "Status: 10%";
                    break;
                case 2:
                    progressBar1.Value = 20;
                    progressBar1.ForeColor = Color.OrangeRed;
                    label12.Text = "Status: 20%";
                    break;
                case 3:
                    progressBar1.Value = 30;
                    progressBar1.ForeColor = Color.DarkOrange;
                    label12.Text = "Status: 30%";
                    break;
                case 4:
                    progressBar1.Value = 40;
                    progressBar1.ForeColor = Color.Orange;
                    label12.Text = "Status: 40%";
                    break;
                case 5:
                    progressBar1.Value = 50;
                    progressBar1.ForeColor = Color.Gold;
                    label12.Text = "Status: 50%";
                    break;
                case 6:
                    System.Drawing.Color someYellowColor = System.Drawing.ColorTranslator.FromHtml("#f4c741");
                    progressBar1.Value = 60;
                    progressBar1.ForeColor = someYellowColor;
                    label12.Text = "Status: 60%";
                    break;
                case 7:
                    System.Drawing.Color someGreenColor = System.Drawing.ColorTranslator.FromHtml("#ebf441");

                    progressBar1.Value = 70;
                    progressBar1.ForeColor = someGreenColor;
                    label12.Text = "Status: 70%";
                    break;
                case 8:
                    progressBar1.Value = 80;
                    progressBar1.ForeColor = Color.GreenYellow;
                    label12.Text = "Status: 80%";
                    break;
                case 9:
                    progressBar1.Value = 90;
                    progressBar1.ForeColor = Color.LawnGreen;
                    label12.Text = "Status: 90%";
                    break;
                case 10:
                    progressBar1.Value = 100;
                    progressBar1.ForeColor = Color.Lime;
                    label12.Text = "Status: 100%";
                    break;
            }

            ExcelApp.Quit();
        }





        public void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            richtextboxTextWasChanged = true;

            

            //string filePath = @"C:\Users\14025\Documents\File Consultants\Brandon\Notes\" + textBox1.Text + "_" + textBox2.Text + "_Notes.rtf";

            //string textToSave = richTextBox1.Rtf;

            

            //SaveAndWriteToRTFFile save = new SaveAndWriteToRTFFile();
            //save.SaveAndWriteToRTFFileMethod(textToSave, filePath);
            ////System.Threading.Thread.Sleep(500);
            //richTextBox1.SaveFile(filePath, RichTextBoxStreamType.RichText);
        }
        private void richTextBox1_LinkClicked(object sender, LinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start(e.LinkText);
        }


       


        public void button1_Click(object sender, EventArgs e)//Add Another Implementation
        {
            Form_AddImplementation form = new Form_AddImplementation();
            form.Show();
            dataGridView1.Rows[dataGridView1.CurrentRow.Index].Selected = true;
        }

        private void button2_Click(object sender, EventArgs e)//Delete Selected Rows
        {

            CheckTextbox1and2 check = new CheckTextbox1and2();
            bool check2 = check.CheckTextbox1and2Method(textBox1.Text, textBox2.Text);

            if (dataGridView1.SelectedRows.ToString() == "")
            {
                dataGridView1.Rows.RemoveAt(dataGridView1.CurrentRow.Index);
            }

            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();

            ExcelApp.Visible = false;

            string excelDBFilePath = @"C:\Users\14025\Documents\File Consultants\ImpList.xls";


            Microsoft.Office.Interop.Excel.Workbook ExcelWorkbook = ExcelApp.Workbooks.Open(excelDBFilePath);

            Microsoft.Office.Interop.Excel.Worksheet ExcelWorksheet = ExcelWorkbook.Sheets[1];


            Microsoft.Office.Interop.Excel.Range colRange = ExcelWorksheet.Columns["B:B"];//get the range object where you want to search from

            string searchString = textBox2.Text;


            Microsoft.Office.Interop.Excel.Range resultRange = colRange.Find(

                What: searchString,

                LookIn: Microsoft.Office.Interop.Excel.XlFindLookIn.xlValues,

                LookAt: Microsoft.Office.Interop.Excel.XlLookAt.xlPart,

                SearchOrder: Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows,

                SearchDirection: Microsoft.Office.Interop.Excel.XlSearchDirection.xlNext

                );// search searchString in the range, if find result, return a range



            if (resultRange == null)

            {

                MessageBox.Show("ERID " + searchString + " not found in database.");

            }
            else
            {

                DialogResult result = MessageBox.Show("Are you sure you want to delete this Implementation?\r\r" + textBox1.Text + " - " + textBox2.Text, "Permanently Delete Implementation?", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                if (result == DialogResult.Yes)
                {

                    //filter textbox1 and 2 to remove carriage returns

                    textBox1.Text = textBox1.Text.TrimEnd('\r', '\n');
                    textBox2.Text = textBox2.Text.TrimEnd('\r', '\n');

                    string directory = @"C:\Users\14025\Documents\File Consultants\Brandon\" + textBox1.Text + "_" + textBox2.Text;
                    //delete employer directory
                    if (Directory.Exists(directory))
                    {
                        System.IO.DirectoryInfo di = new DirectoryInfo(directory);

                        foreach (FileInfo file in di.GetFiles())
                        {
                            file.Delete();
                        }
                        foreach (DirectoryInfo dir in di.GetDirectories())
                        {
                            dir.Delete(true);
                        }
                        Directory.Delete(directory);

                    }

                    //delete Notes file in the Notes folder for this employerID

                    //search ERID in folder
                    SearchForNotes search = new SearchForNotes();
                    string nameToDelete = search.SearchForNotesMethod(textBox2.Text);
                    string fileToDelete = @"C:\Users\14025\Documents\File Consultants\Brandon\Notes\" + nameToDelete;

                    //delete specified file
                    if (check2 == true)
                    {
                        MessageBox.Show("","",MessageBoxButtons.OK,MessageBoxIcon.Information); 
                    }
                    else if(File.Exists(fileToDelete))
                    {

                        File.Delete(fileToDelete);
                        
                    }
                    else if(!File.Exists(fileToDelete))
                    {
                        string nameDoesntExist = search.SearchForNotesMethod("");
                        MessageBox.Show("The Notes file for this employer does not exist.", "No Need To Delete File", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    }

                    // Delete Entire Row - below rows will shift up

                    foreach (DataGridViewRow row in dataGridView1.SelectedRows)
                    {

                        resultRange.EntireRow.Delete(Type.Missing);
                        dataTable1.Rows.RemoveAt(row.Index);
                        ExcelWorkbook.Save();
                        ExcelApp.Quit();


                        if (row.Index == -1)
                        {
                            //dataGridView1.ClearSelection();
                            //dataGridView1.SelectAll();
                            dataGridView1.Rows[dataGridView1.RowCount - 1].Selected = false;
                        }


                    }



                    //int buff = dataGridView1.currentRow.Index;
                }
                else if (result == DialogResult.No)
                {
                    ExcelApp.Quit();
                }


            }
            dataGridView1.Rows[dataGridView1.CurrentRow.Index].Selected = true;

        }

        private void button3_Click(object sender, EventArgs e)//Save notes
        {

            button3WasClicked = true;

            //DateTime date = DateTime.Now;
            string filePath = @"C:\Users\14025\Documents\File Consultants\Brandon\Notes\" + textBox1.Text + "_" + textBox2.Text + "_Notes.rtf";

            string textToSave = richTextBox1.Text;



            SaveAndWriteToRTFFile save = new SaveAndWriteToRTFFile();
            save.SaveAndWriteToRTFFileMethod(textToSave, filePath);

            richTextBox1.SaveFile(filePath, RichTextBoxStreamType.RichText);



            //save to excel spreadsheet
            //string locationColumn = "ERID";
            //string erid = textBox2.Text;
            //string column = "notes";
            //string value = richTextBox1.Text;

            //ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
            //dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);



        }

        private void button4_Click(object sender, EventArgs e)//Clear Tasks

        {
            //clear all previous checks in checklistbox
            foreach (int i in checkedListBox1.CheckedIndices)
            {
                checkedListBox1.SetItemCheckState(i, CheckState.Unchecked);
            }
            progressBar1.Value = 0;
            label12.Text = "Status: 0%";

            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();

            //save checkedListBox1 data to datagridview1

            int currentRow = dataGridView1.CurrentCell.RowIndex;

            if (checkedListBox1.GetItemCheckState(0) == CheckState.Checked)
            {
                dataGridView1[19, currentRow].Value = "1";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx1";
                string value = "1";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }
            else if (checkedListBox1.GetItemCheckState(0) == CheckState.Unchecked)
            {
                dataGridView1[19, currentRow].Value = "0";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx1";
                string value = "0";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }

            if (checkedListBox1.GetItemCheckState(1) == CheckState.Checked)
            {
                dataGridView1[20, currentRow].Value = "1";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx2";
                string value = "1";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }
            else if (checkedListBox1.GetItemCheckState(1) == CheckState.Unchecked)
            {
                dataGridView1[20, currentRow].Value = "0";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx2";
                string value = "0";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }
            if (checkedListBox1.GetItemCheckState(2) == CheckState.Checked)
            {
                dataGridView1[21, currentRow].Value = "1";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx3";
                string value = "1";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }
            else if (checkedListBox1.GetItemCheckState(2) == CheckState.Unchecked)
            {
                dataGridView1[21, currentRow].Value = "0";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx3";
                string value = "0";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }

            if (checkedListBox1.GetItemCheckState(3) == CheckState.Checked)
            {
                dataGridView1[22, currentRow].Value = "1";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx4";
                string value = "1";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }
            else if (checkedListBox1.GetItemCheckState(3) == CheckState.Unchecked)
            {
                dataGridView1[22, currentRow].Value = "0";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx4";
                string value = "0";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }

            if (checkedListBox1.GetItemCheckState(4) == CheckState.Checked)
            {
                dataGridView1[23, currentRow].Value = "1";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx5";
                string value = "1";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }
            else if (checkedListBox1.GetItemCheckState(4) == CheckState.Unchecked)
            {
                dataGridView1[23, currentRow].Value = "0";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx5";
                string value = "0";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }

            if (checkedListBox1.GetItemCheckState(5) == CheckState.Checked)
            {
                dataGridView1[24, currentRow].Value = "1";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx6";
                string value = "1";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }
            else if (checkedListBox1.GetItemCheckState(5) == CheckState.Unchecked)
            {
                dataGridView1[24, currentRow].Value = "0";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx6";
                string value = "0";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }

            if (checkedListBox1.GetItemCheckState(6) == CheckState.Checked)
            {
                dataGridView1[25, currentRow].Value = "1";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx7";
                string value = "1";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }
            else if (checkedListBox1.GetItemCheckState(6) == CheckState.Unchecked)
            {
                dataGridView1[25, currentRow].Value = "0";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx7";
                string value = "0";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }

            if (checkedListBox1.GetItemCheckState(7) == CheckState.Checked)
            {
                dataGridView1[26, currentRow].Value = "1";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx8";
                string value = "1";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }
            else if (checkedListBox1.GetItemCheckState(7) == CheckState.Unchecked)
            {
                dataGridView1[26, currentRow].Value = "0";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx8";
                string value = "0";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }

            if (checkedListBox1.GetItemCheckState(8) == CheckState.Checked)
            {
                dataGridView1[27, currentRow].Value = "1";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx9";
                string value = "1";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }
            else if (checkedListBox1.GetItemCheckState(8) == CheckState.Unchecked)
            {
                dataGridView1[27, currentRow].Value = "0";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx9";
                string value = "0";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }

            if (checkedListBox1.GetItemCheckState(9) == CheckState.Checked)
            {
                dataGridView1[28, currentRow].Value = "1";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx10";
                string value = "1";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }
            else if (checkedListBox1.GetItemCheckState(9) == CheckState.Unchecked)
            {
                dataGridView1[28, currentRow].Value = "0";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx10";
                string value = "0";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }

        }

        private void button5_Click(object sender, EventArgs e)//Completed Implementation
        {

            DialogResult result = MessageBox.Show("Are you sure this Implementation is ready to be completed?\r\r" + textBox1.Text + " - " + textBox2.Text, "Remove and Archive this Implementation?", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

            if (result == DialogResult.Yes)
            {
                //get datagridview data for selected implementation and store in a string array

                DataGridViewRow row = dataGridView1.CurrentRow;
                string stringToAppend = ReturnImplementationAsString.Format(row);



                richTextBox1.AppendText("\r\r" + "Implementation Details:" + "\r\r" + stringToAppend);



                string filePath = @"C:\Users\14025\Documents\File Consultants\Brandon\Notes\" + textBox1.Text + "_" + textBox2.Text + "_Notes.rtf";


                SaveAndWriteToRTFFile save = new SaveAndWriteToRTFFile();
                save.SaveAndWriteToRTFFileMethod(richTextBox1.Rtf, filePath);


                try
                {
                    richTextBox1.SaveFile(filePath, RichTextBoxStreamType.RichText);
                }
                catch (System.IO.IOException ex)
                {
                    MessageBox.Show(ex.ToString());
                }
                catch (System.ArgumentException ex)
                {
                    MessageBox.Show(ex.ToString());
                }

                //move the notes file to ER folder
                string destinationFilePathAndFilename = @"C:\Users\14025\Documents\File Consultants\Brandon\" + textBox1.Text + "_" + textBox2.Text + @"\" + textBox1.Text + "_" + textBox2.Text + "_Notes.rtf";
                MoveFile moveNotesFile = new MoveFile();
                moveNotesFile.MoveFileMethod(filePath, destinationFilePathAndFilename);

                DateTime today = DateTime.Now;
                string todaysDate = today.ToString("yyyyMMdd");
                string sourceDirectoryToZip = @"C:\Users\14025\Documents\File Consultants\Brandon\" + textBox1.Text + "_" + textBox2.Text;
                string destinationZipFileName = @"C:\Users\14025\Documents\File Consultants\Brandon\Archive\" + textBox1.Text + "_" + textBox2.Text + "_"+todaysDate+"_COMPLETED.zip";
                ZipFile zip = new ZipFile();
                zip.ZipFileMethod(sourceDirectoryToZip, destinationZipFileName);

                //delete imp from database

                if (dataGridView1.SelectedRows.ToString() == "")
                {
                    dataGridView1.Rows.RemoveAt(dataGridView1.CurrentRow.Index);
                }

                Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();

                string excelDBFilePath = @"C:\Users\14025\Documents\File Consultants\ImpList.xls";


                Microsoft.Office.Interop.Excel.Workbook ExcelWorkbook = ExcelApp.Workbooks.Open(excelDBFilePath);

                Microsoft.Office.Interop.Excel.Worksheet ExcelWorksheet = ExcelWorkbook.Sheets[1];


                Microsoft.Office.Interop.Excel.Range colRange = ExcelWorksheet.Columns["B:B"];//get the range object where you want to search from

                string searchString = textBox2.Text;


                Microsoft.Office.Interop.Excel.Range resultRange = colRange.Find(

                    What: searchString,

                    LookIn: Microsoft.Office.Interop.Excel.XlFindLookIn.xlValues,

                    LookAt: Microsoft.Office.Interop.Excel.XlLookAt.xlPart,

                    SearchOrder: Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows,

                    SearchDirection: Microsoft.Office.Interop.Excel.XlSearchDirection.xlNext

                    );// search searchString in the range, if find result, return a range



                if (resultRange == null)
                {

                    MessageBox.Show("ERID " + searchString + " not found in database.");

                }
                else
                {

                    string folder = @"C:\Users\14025\Documents\File Consultants\Brandon\" + textBox1.Text + "_" + textBox2.Text;

                   
                    
                    if (Directory.Exists(folder))
                    {
                        //clear employer directory so you can delete it
                        ClearDirectory clear = new ClearDirectory();
                        clear.ClearDirectoryMethod(folder);

                        //delete employer directory
                        try
                        {
                            Directory.Delete(folder);
                        }
                        catch (System.IO.IOException ex)
                        {
                            MessageBox.Show("Method: button5_Click()\r\r" + ex, "Could not delete employer directory", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        catch (System.ArgumentException ex)
                        {
                            MessageBox.Show("Method: button5_Click()\r\r" + ex, "Could not delete employer directory", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                       
                        
                    }


                    // Delete Entire Row - below rows will shift up

                    foreach (DataGridViewRow row1 in dataGridView1.SelectedRows)
                    {

                        resultRange.EntireRow.Delete(Type.Missing);
                        dataTable1.Rows.RemoveAt(row1.Index);
                        ExcelWorkbook.Save();
                        ExcelApp.Quit();


                        if (row1.Index == -1)
                        {
                            //dataGridView1.ClearSelection();
                            //dataGridView1.SelectAll();
                            dataGridView1.Rows[dataGridView1.RowCount - 1].Selected = false;
                        }


                    }
                }
            }

            else if (result == DialogResult.No)
            {

            }

            

            dataGridView1.Rows[dataGridView1.CurrentRow.Index].Selected = true;


        }

        private void button6_Click(object sender, EventArgs e)//Create Email
        {
            //Combo Box List can be found in the ComboBox properties.

            int currentRow = dataGridView1.CurrentRow.Index;

            string internalContactEmail1 = dataGridView1[13, currentRow].Value.ToString();
            string internalContactEmail2 = dataGridView1[33, currentRow].Value.ToString();
            string internalContactEmail3 = dataGridView1[37, currentRow].Value.ToString();
            string internalContactEmail4 = dataGridView1[41, currentRow].Value.ToString();

            string externalContactEmail1 = dataGridView1[17, currentRow].Value.ToString();
            string externalContactEmail2 = dataGridView1[45, currentRow].Value.ToString();
            string externalContactEmail3 = dataGridView1[49, currentRow].Value.ToString();
            string externalContactEmail4 = dataGridView1[53, currentRow].Value.ToString();


            CreateTestResultOutlookEmail email = new CreateTestResultOutlookEmail();

            if (comboBox1.SelectedIndex == 0)
            {
                string emailSubjectLine = textBox1.Text + " - " + textBox2.Text + " - PayFlex Test File Results";


                string emailBody = String.Format("<p style = \"font-size:11pt;\">Hello Everyone,<br/><br/>" +
               "Test results are below:<br/></p> ");

                
                email.CreateTestResultOutlookEmailMethod(externalContactEmail1, externalContactEmail2, externalContactEmail3, externalContactEmail4,
                                               internalContactEmail1, internalContactEmail2, internalContactEmail3, internalContactEmail4,
                                               emailSubjectLine, emailBody);
            }
            else if (comboBox1.SelectedIndex == 1)
            {

                string emailSubjectLine = textBox1.Text + " - " + textBox2.Text + " - PayFlex Test File ETA Request";


                string emailBody = String.Format("<p style = \"font-size:11pt;\">Hello Everyone,<br/><br/>" +
               "Can you please provide a date for which you can send over the next test file to PayFlex for file testing?</p> ");


                email.CreateTestResultOutlookEmailMethod(externalContactEmail1, externalContactEmail2, externalContactEmail3, externalContactEmail4,
                                               internalContactEmail1, internalContactEmail2, internalContactEmail3, internalContactEmail4,
                                               emailSubjectLine, emailBody);
            }
            else if (comboBox1.SelectedIndex == 2)
            {

                string emailSubjectLine = textBox1.Text + " - " + textBox2.Text + " - PayFlex File Testing is Complete";


                string emailBody = String.Format("<p style = \"font-size:11pt;\">Hello Everyone,<br/><br/>" +
               "With the approval of your files in testing, this implementation has moved to production and I will now close it. If you have further file related questions please direct them to the PayFlex account manager "+textBox9.Text+" - "+ internalContactEmail1 + ". Thank you for business and it was a pleasure working with you.</p> ");


                email.CreateTestResultOutlookEmailMethod(externalContactEmail1, externalContactEmail2, externalContactEmail3, externalContactEmail4,
                                               internalContactEmail1, internalContactEmail2, internalContactEmail3, internalContactEmail4,
                                               emailSubjectLine, emailBody);
            }








        }

        private void button7_Click(object sender, EventArgs e)//Check All Tasks
        {
            //check all boxes in checklistbox
            for (int index = 0; index < checkedListBox1.Items.Count; ++index)
            {
                checkedListBox1.SetItemChecked(index, true);
            }
            progressBar1.Value = 100;
            label12.Text = "Status: 100%";

            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();

            //save checkedListBox1 data to datagridview1

            int currentRow = dataGridView1.CurrentCell.RowIndex;

            if (checkedListBox1.GetItemCheckState(0) == CheckState.Checked)
            {
                dataGridView1[19, currentRow].Value = "1";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx1";
                string value = "1";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }
            else if (checkedListBox1.GetItemCheckState(0) == CheckState.Unchecked)
            {
                dataGridView1[19, currentRow].Value = "0";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx1";
                string value = "0";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }

            if (checkedListBox1.GetItemCheckState(1) == CheckState.Checked)
            {
                dataGridView1[20, currentRow].Value = "1";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx2";
                string value = "1";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }
            else if (checkedListBox1.GetItemCheckState(1) == CheckState.Unchecked)
            {
                dataGridView1[20, currentRow].Value = "0";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx2";
                string value = "0";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }
            if (checkedListBox1.GetItemCheckState(2) == CheckState.Checked)
            {
                dataGridView1[21, currentRow].Value = "1";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx3";
                string value = "1";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }
            else if (checkedListBox1.GetItemCheckState(2) == CheckState.Unchecked)
            {
                dataGridView1[21, currentRow].Value = "0";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx3";
                string value = "0";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }

            if (checkedListBox1.GetItemCheckState(3) == CheckState.Checked)
            {
                dataGridView1[22, currentRow].Value = "1";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx4";
                string value = "1";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }
            else if (checkedListBox1.GetItemCheckState(3) == CheckState.Unchecked)
            {
                dataGridView1[22, currentRow].Value = "0";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx4";
                string value = "0";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }

            if (checkedListBox1.GetItemCheckState(4) == CheckState.Checked)
            {
                dataGridView1[23, currentRow].Value = "1";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx5";
                string value = "1";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }
            else if (checkedListBox1.GetItemCheckState(4) == CheckState.Unchecked)
            {
                dataGridView1[23, currentRow].Value = "0";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx5";
                string value = "0";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }

            if (checkedListBox1.GetItemCheckState(5) == CheckState.Checked)
            {
                dataGridView1[24, currentRow].Value = "1";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx6";
                string value = "1";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }
            else if (checkedListBox1.GetItemCheckState(5) == CheckState.Unchecked)
            {
                dataGridView1[24, currentRow].Value = "0";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx6";
                string value = "0";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }

            if (checkedListBox1.GetItemCheckState(6) == CheckState.Checked)
            {
                dataGridView1[25, currentRow].Value = "1";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx7";
                string value = "1";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }
            else if (checkedListBox1.GetItemCheckState(6) == CheckState.Unchecked)
            {
                dataGridView1[25, currentRow].Value = "0";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx7";
                string value = "0";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }

            if (checkedListBox1.GetItemCheckState(7) == CheckState.Checked)
            {
                dataGridView1[26, currentRow].Value = "1";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx8";
                string value = "1";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }
            else if (checkedListBox1.GetItemCheckState(7) == CheckState.Unchecked)
            {
                dataGridView1[26, currentRow].Value = "0";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx8";
                string value = "0";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }

            if (checkedListBox1.GetItemCheckState(8) == CheckState.Checked)
            {
                dataGridView1[27, currentRow].Value = "1";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx9";
                string value = "1";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }
            else if (checkedListBox1.GetItemCheckState(8) == CheckState.Unchecked)
            {
                dataGridView1[27, currentRow].Value = "0";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx9";
                string value = "0";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }

            if (checkedListBox1.GetItemCheckState(9) == CheckState.Checked)
            {
                dataGridView1[28, currentRow].Value = "1";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx10";
                string value = "1";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }
            else if (checkedListBox1.GetItemCheckState(9) == CheckState.Unchecked)
            {
                dataGridView1[28, currentRow].Value = "0";

                string locationColumn = "ERID";
                string erid = textBox2.Text;

                string column = "chckbx10";
                string value = "0";

                ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
                dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);
                ExcelApp.Quit();
            }
        }

        private void button8_Click(object sender, EventArgs e)//Open Files Folder
        {
            Process.Start(@"C:\Users\14025\Documents\File Consultants\Brandon\" + textBox1.Text + "_" + textBox2.Text);

        }

        private void button10_Click(object sender, EventArgs e)//Open Employer Folder
        {
            if (textBox2.Text == "")
            {

            }
            else
            {
                GetGroupName nameAndER = new GetGroupName();
                nameAndER.GetGroupNameMethod(textBox2.Text);
                string NameAndERID = GetGroupName.GroupName;

                Process.Start(@"C:\Users\14025\Documents\File Consultants\Groups\" + NameAndERID);
            }
        }




        private void todaysDateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DateTime today = DateTime.Today;

            richTextBox1.SelectedText = today.ToString("MM/dd/yyyy");
        }

        private void todaysDateAndTimeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DateTime today = DateTime.Now;

            richTextBox1.SelectedText = today.ToString("MM/dd/yyyy - hh:mm:ss");
        }

        private void pFPFMxxxxxxTESTyyyymmddhhmmsstxtToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.SelectedText = "PF.PFM.xxxxxx_TEST_yyyymmdd_hhmmss.txt";
        }

        private void pFPFMxxxxxxEligibilityyyyymmddhhmmsstxtToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.SelectedText = "PF.PFM.xxxxxx_Eligibility_yyyymmdd_hhmmss.txt";
        }

        private void pFPFMxxxxxxDeposityyyymmddhhmmsstxtToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.SelectedText = "PF.PFM.xxxxxx_Deposit_yyyymmdd_hhmmss.txt";
        }

        private void highlightSelectedTextToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.SelectionBackColor = System.Drawing.Color.Yellow;
        }


        private void undoToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void boldToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.SelectionFont = new System.Drawing.Font(richTextBox1.Font, FontStyle.Bold);
        }

        private void italicsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.SelectionFont = new System.Drawing.Font(richTextBox1.Font, FontStyle.Italic);
        }

        private void underlineToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.SelectionFont = new System.Drawing.Font(richTextBox1.Font, FontStyle.Underline);
        }

        private void strikeoutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.SelectionFont = new System.Drawing.Font(richTextBox1.Font, FontStyle.Strikeout);
        }
        private void normalTextToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.SelectionFont = new System.Drawing.Font(richTextBox1.Font, FontStyle.Regular);
        }
        private void requiredNotesToolStripMenuItem_Click(object sender, EventArgs e)
        {

            int currentRow = dataGridView1.CurrentCell.RowIndex;
            int currentColumn = dataGridView1.CurrentCell.ColumnIndex;


            richTextBox1.SelectedText = "Summary Notes\r\n";
            richTextBox1.SelectedText = "__________________________________\r\n";
            richTextBox1.SelectedText = "Name: "+textBox1.Text+"\r\n";
            richTextBox1.SelectedText = "ERID: " + textBox2.Text + "\r\n";
            richTextBox1.SelectedText = "Current Products: " + textBox6.Text + "\r\n";
            richTextBox1.SelectedText = "Products being Added: " + textBox7.Text + "\r\n";
            richTextBox1.SelectedText = "Plan Year Effective Date: " + textBox5.Text + "\r\n";
            richTextBox1.SelectedText = "Implementation Deadline: " + textBox10.Text + "\r\n\r\n";
            richTextBox1.SelectedText = "AM:  " + dataGridView1.Rows[currentRow].Cells[11].Value.ToString() + " - "+ dataGridView1.Rows[currentRow].Cells[13].Value.ToString() + "\r\n";
            richTextBox1.SelectedText = "IM: " + dataGridView1.Rows[currentRow].Cells[30].Value.ToString() + "\r\r\n";

            richTextBox1.SelectedText = "File Sender Contact Name: " + dataGridView1.Rows[currentRow].Cells[15].Value.ToString() + "\r\n";
            richTextBox1.SelectedText = "File Sender Contact Email: " + dataGridView1.Rows[currentRow].Cells[17].Value.ToString() + "\r\n";
            richTextBox1.SelectedText = "File Sender Contact Type: " + dataGridView1.Rows[currentRow].Cells[18].Value.ToString() + "\r\n";
            richTextBox1.SelectedText = "File Type being sent: " + dataGridView1.Rows[currentRow].Cells[19].Value.ToString() + "\r\r\n";

            richTextBox1.SelectedText = "File Sender Contact Name: " + dataGridView1.Rows[currentRow].Cells[15].Value.ToString() + "\r\n";
            richTextBox1.SelectedText = "File Sender Contact Email: " + dataGridView1.Rows[currentRow].Cells[17].Value.ToString() + "\r\n";
            richTextBox1.SelectedText = "File Sender Contact Type: " + dataGridView1.Rows[currentRow].Cells[18].Value.ToString() + "\r\n";
            richTextBox1.SelectedText = "File Type being sent: " + dataGridView1.Rows[currentRow].Cells[19].Value.ToString() + "\r\r\n";

            richTextBox1.SelectedText = "ETL Ticket Link: \r\r\n";
            richTextBox1.SelectedText = "HSA: \r\n";
            richTextBox1.SelectedText = "Debit cards: \r\n";
            richTextBox1.SelectedText = "Wired Commute Benefits: \r\n";
            richTextBox1.SelectedText = "Division Codes: \r\n";
            richTextBox1.SelectedText = "Payroll Schedule ID(s): \r\r\n";
            richTextBox1.SelectedText = "HSA CSA's: \r\n";
            richTextBox1.SelectedText = "FSA CSA's: \r\r\n";
            richTextBox1.SelectedText = "Aetna Files Currently Being Sent(File Sender, File Type): \r\n";
            richTextBox1.SelectedText = "__________________________________\r\n\r\n";
        }



        private void changeBackgroundColorToolStripMenuItem1_Click(object sender, EventArgs e)
        {

            int currentRow = dataGridView1.CurrentRow.Index;
            ColorDialog changeBackgroundColor = new ColorDialog();
            if (changeBackgroundColor.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                richTextBox1.BackColor = changeBackgroundColor.Color;

            }
        }

        private void changeTextColorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ColorDialog changeTextColor = new ColorDialog();
            if (changeTextColor.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                richTextBox1.ForeColor = changeTextColor.Color;
            }
        }
                                    


        private void dataGridView1_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;
        }

        private void dataGridView1_DragDrop(object sender, DragEventArgs e)
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
                   region, segment,
                   benefitEffectiveDate, currentProducts,
                    addedProducts, newImpFlag,
                   IM_AM, impDeadline, sftpFlag, contactName,
                     contactphoneNumber, contactEmail, contactType,
                     fileLayout;

                        GetWordFileData passInputToMethod = new GetWordFileData();//create instance variable

                        //pull word data off of the FileConsultantRequestForm.docx into this form
                        passInputToMethod.GetWordFileDataMethod(file, out employerName, out employerID,
                                                                out region, out segment,
                                                                out benefitEffectiveDate, out currentProducts,
                                                                out addedProducts, out newImpFlag,
                                                                out IM_AM, out impDeadline, out sftpFlag, out contactName,
                                                                out contactphoneNumber, out contactEmail, out contactType,
                                                                out fileLayout);


                        foreach (DataGridViewRow row in dataGridView1.Rows)
                        {
                            if (row.Cells[2].Value.ToString().Equals(employerID))
                            {
                                MessageBox.Show("This ERID already exists in this list. Only one implementation per employer.", "No-No! Bad File Consultant!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                break;
                            }
                        }
                        Regex rg = new Regex("[A-Za-z0-9/ .]^");

                        bool thing = rg.IsMatch("•");
                        if (thing == true)
                        {
                            employerName = employerName + rg.Replace("•", "");
                            employerID = employerID + rg.Replace("•", "");
                        }
                        

                        //save data to database
                        ExcelDataBasePush db = new ExcelDataBasePush();
                        db.ExcelDataBasePushMethod(employerName, employerID, region, segment,
                                                benefitEffectiveDate, currentProducts, addedProducts, newImpFlag,
                                                IM_AM, impDeadline, sftpFlag,
                                                string.Empty, string.Empty, string.Empty, string.Empty,
                                                contactName, contactphoneNumber,
                                                contactEmail, contactType, fileLayout,
                                                string.Empty, string.Empty, string.Empty, string.Empty,
                                                string.Empty, string.Empty, string.Empty, string.Empty,
                                                string.Empty, string.Empty, string.Empty, string.Empty,
                                                string.Empty, string.Empty, string.Empty, string.Empty,
                                                string.Empty, string.Empty, string.Empty, string.Empty,
                                                string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty,
                                                string.Empty, string.Empty, string.Empty, string.Empty,
                                                string.Empty, string.Empty, string.Empty, string.Empty);


                        textBox1.Text = employerName;
                        textBox2.Text = employerID;

                        //filter textbox1 and 2 to remove carriage returns

                        textBox1.Text = textBox1.Text.TrimEnd('\r', '\n');
                        textBox2.Text = textBox2.Text.TrimEnd('\r', '\n');

                        //create employer directory
                         Directory.CreateDirectory(@"C:\Users\14025\Documents\File Consultants\Brandon\" + employerName + "_" + employerID);

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
            catch (System.Exception ex)
            {

                MessageBox.Show("Method: AddImplementation\r Something prevented data from pulling from the FileConsultantRequestForm.docx.\r\r" + ex);
            }

            Form_OpenImplementationList form = new Form_OpenImplementationList();
            form.Show();

            this.Hide();
        }




 




        private void listView1_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            int counter = 0;
            //ImageList ImageList1 = new ImageList ();
            //string[] filesInDir = Directory.GetFiles(@"C:\Users\14025\Documents\File Consultants\Brandon\" + textBox1.Text + "_" + textBox2.Text);

            //foreach (string file in filesInDir)
            //{

            //    Icon stuff = Icon.ExtractAssociatedIcon(file);
            //    ImageList1.Images.Add(stuff);

            //}

            foreach (string completeFilePath in files)
            {             


                counter++;
                string filePath = System.IO.Path.GetDirectoryName(completeFilePath);
                string fileName = System.IO.Path.GetFileName(completeFilePath);
                string destinationPath = @"C:\Users\14025\Documents\File Consultants\Brandon\" + textBox1.Text + "_" + textBox2.Text + @"\" + fileName;


                //ImageList1.Images.Add(Icon.ExtractAssociatedIcon(fileName))
                //System.Drawing.Icon.ExtractAssociatedIcon(completeFilePath);


                if (!File.Exists(destinationPath))
                {
                    File.Copy(completeFilePath, destinationPath);
                    listView1.Items.Add(fileName);

                }
                else
                {
                    MessageBox.Show("This file already exists in this this employers folder.", "Can't Do It. Wouldn't be Proper.", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }

            }
            label16.Text = "Files: " + counter;
        }

        private void listView1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;//copy the file
            }


        }

        private void listView1_DragOver(object sender, DragEventArgs e)
        {          

           
            

           

                
                //string filePath = System.IO.Path.GetDirectoryName(completeFilePath);
                //string fileName = System.IO.Path.GetFileName(completeFilePath);
                //string destinationPath = @"C:\Users\14025\Documents\File Consultants\Brandon\" + textBox1.Text + "_" + textBox2.Text + @"\" + fileName;


            
        }
        private void listView1_DragLeave(object sender, EventArgs e)
        {
            object folder = e.GetType();

            
                MessageBox.Show(folder.ToString());
            
        }
        private void listView1_ItemDrag(object sender, ItemDragEventArgs e)
        {
            List<object> selection = new List<object>();//create a list

            foreach (ListViewItem item in listView1.SelectedItems)//add the filenames to list
            {
                int imgIndex = item.ImageIndex;
                selection.Add(item);
            }

            DataObject data = new DataObject(DataFormats.FileDrop, selection.ToArray());
            DoDragDrop(data, DragDropEffects.Copy);

 
        }

        private void listView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            string selectedFile = listView1.SelectedItems[0].Text;
            string fileToOpen = @"C:\Users\14025\Documents\File Consultants\Brandon\" + textBox1.Text + "_" + textBox2.Text + @"\" + selectedFile;


            // If it's a file open it
            if (File.Exists(fileToOpen))
            {
                //MessageBox.Show(currentDir + @"\" + selectedFile);
                try
                {
                    System.Diagnostics.Process.Start(fileToOpen);
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.StackTrace);
                }
            }
        }

        private void listView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete)
            {
                var selectedItems = listView1.SelectedItems;


                foreach (ListViewItem selectedItem in selectedItems)
                {
                    listView1.Items.Remove(selectedItem);

                    string stuff = selectedItem.Text;

                    string fileToDelete = @"C:\Users\14025\Documents\File Consultants\Brandon\" + textBox1.Text + "_" + textBox2.Text + @"\" + stuff;


                    if (File.Exists(fileToDelete))
                    {
                        try
                        {
                            File.Delete(fileToDelete);
                        }
                        catch (System.Exception ex)
                        {
                            MessageBox.Show(ex.StackTrace);
                        }

                    }

                }
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Control && e.KeyCode == Keys.C)
            {
                var selectedItems = listView1.SelectedItems;


                foreach (ListViewItem selectedItem in selectedItems)
                {
                    

                    string stuff = selectedItem.Text;

                    string fileToCopy = @"C:\Users\14025\Documents\File Consultants\Brandon\" + textBox1.Text + "_" + textBox2.Text + @"\" + stuff;
                    

                    if (File.Exists(fileToCopy))
                    {
                        try
                        {
                            // Retrieves data  
                            IDataObject iData = Clipboard.GetDataObject();
                            
                        }
                       
                        catch (System.Threading.ThreadStateException ex)
                        {
                            MessageBox.Show(ex.ToString());
                        }
                        catch (System.Runtime.InteropServices.ExternalException ex)
                        {
                            MessageBox.Show(ex.ToString());
                        }

                    }
                }
            }
            else if (e.KeyCode == Keys.Control && e.KeyCode == Keys.V)
            {
                
                try
                {
                    //// Retrieves data  
                    //IDataObject iData = Clipboard.GetDataObject();
                    //// Is Data Text?  
                    //if (iData.GetDataPresent(DataFormats.Text))
                    //    label1.Text = (String)iData.GetData(DataFormats.Text);
                    //else
                    //    label1.Text = "Data not found.";

                    ////Clear method removes all data from the Clipboard.  
                    //Clipboard.Clear();
                }
                catch (System.ArgumentException ex)
                {
                    MessageBox.Show(ex.ToString());
                }
                catch (System.Threading.ThreadStateException ex)
                {
                    MessageBox.Show(ex.ToString());
                }
                catch (System.Runtime.InteropServices.ExternalException ex)
                {
                    MessageBox.Show(ex.ToString());
                }

                
                    
                
            }


        }



        //Listview context menu

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void cutToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void pasteToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void deleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var selectedItems = listView1.SelectedItems;


            foreach (ListViewItem selectedItem in selectedItems)
            {
                listView1.Items.Remove(selectedItem);

                string stuff = selectedItem.Text;

                string fileToDelete = @"C:\Users\14025\Documents\File Consultants\Brandon\" + textBox1.Text + "_" + textBox2.Text + @"\" + stuff;

                

                if (File.Exists(fileToDelete))
                {
                    try
                    {
                        File.Delete(fileToDelete);
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show(ex.StackTrace);
                    }

                }




            }

           
        }

        private void addToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void changeEmployerNameToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string originalName = textBox1.Text;

            Directory.Move(@"C:\Users\14025\Documents\File Consultants\Brandon\" + originalName + "_" + textBox2.Text,
                           @"C:\Users\14025\Documents\File Consultants\Brandon\" + textBox1.Text + "_" + textBox2.Text);

            File.Move(@"C:\Users\14025\Documents\File Consultants\Brandon\Notes\" + originalName + "_" + textBox2.Text + "_Notes.rtf",
                           @"C:\Users\14025\Documents\File Consultants\Brandon\Notes\" + textBox1.Text + "_" + textBox2.Text + "_Notes.rtf");

        }
        private void eRNameERIDToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.SelectedText = textBox1.Text+" - "+textBox2.Text;

        }

        private void textBox12_TextChanged(object sender, EventArgs e)
        {
            if (richTextBox1.Text.Length > 0)
            {

            

                int startPos = FindMyTextPosition(textBox12.Text, 0 , richTextBox1.Text.Length);
                try
                {
                    richTextBox1.Select(startPos, textBox12.Text.Length);
                    richTextBox1.SelectionStart = startPos;
                    richTextBox1.Focus();
                    richTextBox1.SelectionBackColor = Color.Goldenrod;
                    textBox12.Focus();

                    

                }
                catch (ArgumentOutOfRangeException ex)
                {
                    if (textBox12.Text.Length <= 0)
                    {
                        richTextBox1.SelectAll();
                        richTextBox1.Focus();
                        richTextBox1.SelectionBackColor = Color.White;
                        richTextBox1.Select(0,0);
                        richTextBox1.SelectionStart = 0;
                        textBox12.Select();

                        //System.Drawing.Point point = new System.Drawing.Point(0, 0);
                        //richTextBox1.PointToScreen(point);
                        //richTextBox1.
                    }
                }
                
                
            }
        }//richtextbox search bar

        public int FindMyTextPosition(string searchText, int searchStart, int searchEnd)
        {
            // Initialize the return value to false by default.
            int returnValue = -1;

            // Ensure that a search string and a valid starting point are specified.
            if (searchText.Length > 0 && searchStart >= 0)
            {
                // Ensure that a valid ending value is provided.
                if (searchEnd > searchStart || searchEnd == -1)
                {
                    // Obtain the location of the search string in richTextBox1.
                    int indexToText = richTextBox1.Find(searchText, searchStart, searchEnd, RichTextBoxFinds.WholeWord);
                    // Determine whether the text was found in richTextBox1.
                    if (indexToText >= 0)
                    {
                        // Return the index to the specified search text.
                        returnValue = indexToText;
                    }
                }
            }

            return returnValue;
        }

        private void newLogEntryToolStripMenuItem_Click(object sender, EventArgs e)
        {

            Form_NewLogEntryInOpenImplementationNotes form_NewLogEntry = new Form_NewLogEntryInOpenImplementationNotes();
            form_NewLogEntry.Show();
            
        }

        public void GetNewLogEntryDataAndDisplay()
        {
            string ERID = Form_NewLogEntryInOpenImplementationNotes.ERID;
            string todaysDate = Form_NewLogEntryInOpenImplementationNotes.TodaysDate;
            string entryType = Form_NewLogEntryInOpenImplementationNotes.EntryType;
            string regarding = Form_NewLogEntryInOpenImplementationNotes.Regarding;


            Clipboard.SetText(todaysDate + " - " + entryType + " - " + regarding );


            //MessageBox.Show("ERID: "+ERID);
            //MessageBox.Show("otherStuff: " + otherStuff);

            //this.Activate();
            //this.richTextBox1.Select();
            //this.richTextBox1.AppendText(ERID +" - "+otherStuff);
        }

        private void Form_OpenImplementationList_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Control && e.KeyCode == Keys.N)
            {
                Form_NewLogEntryInOpenImplementationNotes form = new Form_NewLogEntryInOpenImplementationNotes();
                form.Show();
            }
        }

        private void button9_Click(object sender, EventArgs e)//UnArchive Implementation
        {
            Form_UnArchiveImplementation form = new Form_UnArchiveImplementation();
            form.Show();
        }

        private void button11_Click(object sender, EventArgs e)//add selected
        {
            DuplicateForm = true;

            EmployerName = "";
            EmployerID = "";
            Region = dataGridView1.SelectedCells[2].Value.ToString();
            Segment = dataGridView1.SelectedCells[3].Value.ToString();
            EffectiveDate = dataGridView1.SelectedCells[4].Value.ToString();
            CurrentProducts = dataGridView1.SelectedCells[5].Value.ToString();
            AddingProduct = dataGridView1.SelectedCells[6].Value.ToString();
            NewImp = dataGridView1.SelectedCells[7].Value.ToString();
            AMIMInvolved = dataGridView1.SelectedCells[8].Value.ToString();
            ImpDeadline = dataGridView1.SelectedCells[9].Value.ToString();
            SFTPCreds = dataGridView1.SelectedCells[10].Value.ToString();


            InternalContactName1 = dataGridView1.SelectedCells[11].Value.ToString();
            InternalContactPhone1 = dataGridView1.SelectedCells[12].Value.ToString();
            InternalContactEmail1 = dataGridView1.SelectedCells[13].Value.ToString();
            InternalContactType1 = dataGridView1.SelectedCells[14].Value.ToString();

            ExternalContactName1 = dataGridView1.SelectedCells[15].Value.ToString();
            ExternalContactPhone1 = dataGridView1.SelectedCells[16].Value.ToString();
            ExternalContactEmail1 = dataGridView1.SelectedCells[17].Value.ToString();
            ExternalContactType1 = dataGridView1.SelectedCells[18].Value.ToString();

            InternalContactName2 = dataGridView1.SelectedCells[31].Value.ToString();
            InternalContactPhone2 = dataGridView1.SelectedCells[32].Value.ToString();
            InternalContactEmail2 = dataGridView1.SelectedCells[33].Value.ToString();
            InternalContactType2 = dataGridView1.SelectedCells[34].Value.ToString();

            InternalContactName3 = dataGridView1.SelectedCells[35].Value.ToString();
            InternalContactPhone3 = dataGridView1.SelectedCells[36].Value.ToString();
            InternalContactEmail3 = dataGridView1.SelectedCells[37].Value.ToString();
            InternalContactType3 = dataGridView1.SelectedCells[38].Value.ToString();

            InternalContactName4 = dataGridView1.SelectedCells[39].Value.ToString();
            InternalContactPhone4 = dataGridView1.SelectedCells[40].Value.ToString();
            InternalContactEmail4 = dataGridView1.SelectedCells[41].Value.ToString();
            InternalContactType4 = dataGridView1.SelectedCells[42].Value.ToString();

            ExternalContactName2 = dataGridView1.SelectedCells[43].Value.ToString();
            ExternalContactPhone2 = dataGridView1.SelectedCells[44].Value.ToString();
            ExternalContactEmail2 = dataGridView1.SelectedCells[45].Value.ToString();
            ExternalContactType2 = dataGridView1.SelectedCells[46].Value.ToString();

            ExternalContactName3 = dataGridView1.SelectedCells[47].Value.ToString();
            ExternalContactPhone3 = dataGridView1.SelectedCells[48].Value.ToString();
            ExternalContactEmail3 = dataGridView1.SelectedCells[49].Value.ToString();
            ExternalContactType3 = dataGridView1.SelectedCells[50].Value.ToString();

            ExternalContactName4 = dataGridView1.SelectedCells[51].Value.ToString();
            ExternalContactPhone4 = dataGridView1.SelectedCells[52].Value.ToString();
            ExternalContactEmail4 = dataGridView1.SelectedCells[53].Value.ToString();
            ExternalContactType4 = dataGridView1.SelectedCells[54].Value.ToString();

            Form_AddImplementation form = new Form_AddImplementation();
            form.Show();




       
    }



        //private void dataGridView2_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        //{
        //    Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();

        //    if (dataGridView2.CurrentRow == null && dataGridView2.CurrentRow.Index < dataGridView2.Rows.Count - 1)
        //    {
        //        int currentRow = dataGridView2.CurrentCell.RowIndex + 1;
        //        // Update the DB with new data.
        //        int currentColumn = dataGridView2.CurrentCell.ColumnIndex;
        //        string column = "";

        //        string locationColumn = "ERID";
        //        string erid = textBox2.Text;

        //        if (currentColumn == 0 && currentRow == 0)
        //        {


        //            column = "InternalContactName";
        //            string value = "";

        //            value = dataGridView2.Rows[currentRow].Cells[currentColumn].Value.ToString();

        //            ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
        //            dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);



        //        }
        //        else if (currentColumn == 1 && currentRow == 0)
        //        {
        //            column = "InternalContactEmail";
        //            string value = "";
        //            value = dataGridView2.Rows[currentRow].Cells[currentColumn].Value.ToString();

        //            ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
        //            dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);

        //        }
        //        else if (currentColumn == 2 && currentRow == 0)
        //        {
        //            column = "InternalContactPhone";
        //            string value = "";
        //            value = dataGridView2.Rows[currentRow].Cells[currentColumn].Value.ToString();

        //            ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
        //            dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);

        //        }
        //        else if (currentColumn == 3 && currentRow == 0)
        //        {
        //            column = "InternalContactType";
        //            string value = "";
        //            value = dataGridView2.Rows[currentRow].Cells[currentColumn].Value.ToString();

        //            ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
        //            dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);

        //        }


        //    }
        //    else if (dataGridView2.CurrentRow != null)
        //    {
        //        int currentRow = dataGridView2.CurrentCell.RowIndex;
        //        // Update the DB with new data.
        //        int currentColumn = dataGridView2.CurrentCell.ColumnIndex;
        //        string column = "";

        //        string locationColumn = "ERID";
        //        string erid = textBox2.Text;

        //        if (currentColumn == 0 && currentRow == 0)
        //        {


        //            column = "InternalContactName";
        //            string value = "";

        //            value = dataGridView2.Rows[currentRow].Cells[currentColumn].Value.ToString();

        //            ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
        //            dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);



        //        }
        //        else if (currentColumn == 1 && currentRow == 0)
        //        {
        //            column = "InternalContactEmail";
        //            string value = "";
        //            value = dataGridView2.Rows[currentRow].Cells[currentColumn].Value.ToString();

        //            ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
        //            dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);

        //        }
        //        else if (currentColumn == 2 && currentRow == 0)
        //        {
        //            column = "InternalContactPhone";
        //            string value = "";
        //            value = dataGridView2.Rows[currentRow].Cells[currentColumn].Value.ToString();

        //            ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
        //            dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);

        //        }
        //        else if (currentColumn == 3 && currentRow == 0)
        //        {
        //            column = "InternalContactType";
        //            string value = "";
        //            value = dataGridView2.Rows[currentRow].Cells[currentColumn].Value.ToString();

        //            ExcelDataBaseUpdate dataPush = new ExcelDataBaseUpdate();
        //            dataPush.ExcelDataBaseUpdateMethod(locationColumn, erid, column, value);

        //        }

        //    }








        //}

        //private void dataGridView3_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        //{

        //}



    }

    public static class ReturnImplementationAsString
    {
        public static string Format(this DataGridViewRow row)
        {
            string[] values = new string[row.Cells.Count];
            for (int i = 0; i < row.Cells.Count; i++)
                values[i] = row.Cells[i].Value + "";
            return string.Join("|", values);

            
            
        }
    }


}
