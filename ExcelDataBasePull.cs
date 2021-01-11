using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Data.OleDb;

namespace FCHelper_v001
{


    class ExcelDataBasePull
    {
        private string windowsUserName = System.Environment.UserName;//gives windows username
        private OleDbDataAdapter fillFromSheet1;
        public void ExcelDataBasePullMethod(out OleDbDataAdapter output)//inputs to method
        {

           
            OleDbDataAdapter FillFromSheet = new OleDbDataAdapter();
            output = fillFromSheet1;

            try
            {
                string excelDBFilePath = "C:\\Users\\" + windowsUserName + "\\Documents\\ImpList.xls";
                string excelTableSource = "Provider = Microsoft.Jet.OLEDB.4.0; Data Source = " + excelDBFilePath + "; Extended Properties = 'Excel 8.0; HDR = YES'";

                OleDbConnection connection = new OleDbConnection();
                OleDbCommand command = new OleDbCommand();

                connection = new OleDbConnection(excelTableSource);

                string sql = "SELECT * FROM [sheet1$]";
                fillFromSheet1 = new OleDbDataAdapter(sql, connection);//pulls data from Sheet1 in excel doc to use Fill method                         
                

               

                connection.Open();
                command.Connection = connection;

                command.CommandText = sql;
                command.ExecuteNonQuery();
                FillFromSheet = fillFromSheet1;
                connection.Close();
                sql = null;

                
               

            }
            catch (Exception ex)
            {
                MessageBox.Show("Method: OpenImplementationList\rSomething prevented the data from pulling to the datagridview control.\r\r" + ex, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            

            /*

                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();//)Marshal.GetActiveObject("Excel.Application");
            

            Workbook wb = excel.Workbooks.Open(excelDBFilePath);
            Worksheet ws = wb.ActiveSheet;

            string excelTableSource = "Provider = Microsoft.Jet.OLEDB.4.0; Data Source = " + excelDBFilePath + "; Extended Properties = 'Excel 8.0; HDR = YES;'";

            OleDbConnection connection = new OleDbConnection(excelTableSource);
            OleDbDataAdapter fillFromSheet1 = new OleDbDataAdapter("SELECT * FROM [ImpList$]", connection);//pulls data from Sheet1 in excel doc to use Fill method
            System.Data.DataTable excelTable = new System.Data.DataTable();
            fillFromSheet1.Fill(excelTable);
            dataGridView1.DataSource = excelTable;*/
        }
    }

       
}
