using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using System.Windows.Forms;

namespace FCHelper_v001
{
    class ExcelDataBasePush
    {
        private string windowsUserName = System.Environment.UserName;//gives windows username
        //public OleDbConnection connection = new OleDbConnection();

        public void ExcelDataBasePushMethod(string ername, string erid, string region,
                                        string segment, string effDate, string curProd,
                                        string addProd, string newImp, string AM_IM,
                                        string impDdline, string sftpFlag, 
                                        string inConName, string inConPhone, string inConEmail, string inConType, 
                                        string exConName, string exConPhone, string exConEmail, string exConType, string fileType,
                                        string inConName2, string inConPhone2, string inConEmail2, string inConType2,
                                        string inConName3, string inConPhone3, string inConEmail3, string inConType3,
                                        string inConName4, string inConPhone4, string inConEmail4, string inConType4,
                                        string exConName2, string exConPhone2, string exConEmail2, string exConType2,
                                        string exConName3, string exConPhone3, string exConEmail3, string exConType3,
                                        string exConName4, string exConPhone4, string exConEmail4, string exConType4,
                                        string chckbox1, string chckbox2, string chckbox3, string chckbox4,
                                        string chckbox5, string chckbox6, string chckbox7, string chckbox8,
                                        string chckbox9, string chckbox10)//inputs to method
        {
            string excelDBFilePath = @"C:\Users\14025\Documents\File Consultants\ImpList.xls";
            OleDbConnection connection = new OleDbConnection();
            try
            {
                string excelTableSource = "Provider = Microsoft.Jet.OLEDB.4.0; Data Source = " + excelDBFilePath + "; Extended Properties = 'Excel 8.0; HDR = YES'";

                OleDbCommand command = new OleDbCommand();
                string sql = null;

                connection = new OleDbConnection(excelTableSource);

                connection.Open();
                command.Connection = connection;

                sql = "INSERT INTO [Brandon$] (ERname, ERID, Region, Segment," +//insert into already existing fields, inside the spreadsheet
                    "EffDate, CurrentProduct, AddingProduct," +
                    "NewImplementation, AM_IM, ImplementationDeadline, SFTPFlag," +
                    "InternalContactName, InternalContactPhone, InternalContactEmail, InternalContactType," +
                    "ExternalContactName, ExternalContactPhone, ExternalContactEmail, ExternalContactType, FileType," +

                    "chckbx1,chckbx2,chckbx3,chckbx4,chckbx5,chckbx6,chckbx7,chckbx8,chckbx9,chckbx10," +

                   "InternalContactName2, InternalContactPhone2, InternalContactEmail2, InternalContactType2," +
                   "InternalContactName3, InternalContactPhone3, InternalContactEmail3, InternalContactType3," +
                   "InternalContactName4, InternalContactPhone4, InternalContactEmail4, InternalContactType4," +

                   "ExternalContactName2, ExternalContactPhone2, ExternalContactEmail2, ExternalContactType2, " +
                   "ExternalContactName3, ExternalContactPhone3, ExternalContactEmail3, ExternalContactType3, " +
                   "ExternalContactName4, ExternalContactPhone4, ExternalContactEmail4, ExternalContactType4)" +




                                      "VALUES ('" + ername + "','" + erid + "','" + region +
                                              "','" + segment + "','" + effDate + "','" + curProd +
                                              "','" + addProd + "','" + newImp + "','" + AM_IM + 
                                              "','" + impDdline + "','" + sftpFlag + 
                                              "','" + inConName + "','" + inConPhone + "','" + inConEmail + "','" + inConType +
                                              "','" + exConName + "','" + exConPhone + "','" + exConEmail + "','" + exConType + "','" + fileType +

                 "','" + chckbox1 + "','" + chckbox2 + "','" + chckbox3 +
                 "','" + chckbox4 + "','" + chckbox5 + "','" + chckbox6 +
                 "','" + chckbox7 + "','" + chckbox8 + "','" + chckbox9 +
                  "','" + chckbox10 + 
                                              "','" + inConName2 + "','" + inConPhone2 + "','" + inConEmail2 + "','" + inConType2 +
                                              "','" + inConName3 + "','" + inConPhone3 + "','" + inConEmail3 + "','" + inConType3 +
                                              "','" + inConName4 + "','" + inConPhone4 + "','" + inConEmail4 + "','" + inConType4 +
                                              "','" + exConName2 + "','" + exConPhone2 + "','" + exConEmail2 + "','" + exConType2 +
                                              "','" + exConName3 + "','" + exConPhone3 + "','" + exConEmail3 + "','" + exConType3 +
                                              "','" + exConName4 + "','" + exConPhone4 + "','" + exConEmail4 + "','" + exConType4 + "')";




                command.CommandText = sql;
                command.ExecuteNonQuery();
                connection.Close();
                sql = null;
            }
            catch(Exception ex)
            {
                MessageBox.Show("Something prevented the data from pushing to the Excel database.\r\r" + ex, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


           

          


            /* string filePath = "C:\\Users\\" + userName + "\\Documents\\ImpList.xls";

            Workbook wb = excel.Workbooks.Open(filePath);
            Worksheet ws = wb.ActiveSheet;

            string excelTableSource = "Provider = Microsoft.Jet.OLEDB.4.0; Data Source = " + filePath + "; Extended Properties =\"Excel 8.0; HDR = No;\";";

            OleDbConnection connection = new OleDbConnection(excelTableSource);
            OleDbDataAdapter fillFromSheet1 = new OleDbDataAdapter("SELECT * FROM [ImpList$]", connection);//pulls data from Sheet1 in excel doc to use Fill method
            System.Data.DataTable excelTable = new System.Data.DataTable();
            fillFromSheet1.Fill(excelTable);
            dataGridView1.DataSource = excelTable;*/
        }
    }
}
