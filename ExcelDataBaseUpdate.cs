using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Windows.Forms;

namespace FCHelper_v001
{
    class ExcelDataBaseUpdate
    {
        private string windowsUserName = System.Environment.UserName;//gives windows username


        public void ExcelDataBaseUpdateMethod(string locationColumn, string erid, string columnName,string value)//inputs to method
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
                //UPDATE table SET column = rowvalue WHERE columnToChange = value
                sql = "UPDATE [Brandon$] SET " + columnName + " = @columnName WHERE "+ locationColumn + " = @locationColumn;";
                
                command.Parameters.AddWithValue("@columnName", value);
                command.Parameters.AddWithValue("@locationColumn", erid);

                command.CommandText = sql;
                command.ExecuteNonQuery();
                connection.Close();
                sql = null;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Something prevented the data from pushing to the Excel database.\r\r" + ex, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }







            
        }
    
    }
}
