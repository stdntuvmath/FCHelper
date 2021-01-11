using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace FCHelper_v001
{
    class UpdateDOCSFolderSpreadsheet
    {
        public void UpdateDOCSFolderSpreadsheetMethod(string groupName)
        {
            PrivateUpdateDOCSFolderSpreadsheetMethod(groupName);
        }

        private void PrivateUpdateDOCSFolderSpreadsheetMethod(string groupName)
        {
            string impSpreadsheetDOCS = @"\\phx-fs-02.payflex.com\Data\PFS\GFP\group\"+groupName+@"\DOCS";
            string impSpreadsheetDocs = @"\\phx-fs-02.payflex.com\Data\PFS\GFP\group\" + groupName + @"\Docs";
            string impSpreadsheetdocs = @"\\phx-fs-02.payflex.com\Data\PFS\GFP\group\" + groupName + @"\docs";
            string testPath = @"\\phx-fs-02.payflex.com\GDrive\DataServicesGroup\Personal\Brandon\Brandon's Staging Folder\FCHelperTesting\DOCS\ImpNotes.xls";

            object misValue = System.Reflection.Missing.Value;



            try
            {
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;
                Microsoft.Office.Interop.Excel.Workbook workBook = new Microsoft.Office.Interop.Excel.Workbook();
                Microsoft.Office.Interop.Excel.Worksheet workSheet = new Microsoft.Office.Interop.Excel.Worksheet();

                workBook = excel.Workbooks.Add(misValue);
                workSheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.Worksheets.get_Item(1);

                //insert columns into spreadsheets
                workSheet.Name = "Brandon";
                workSheet.Cells[1, 1] = "Date";
                workSheet.Cells[1, 1] = "Employer ID";
                workSheet.Cells[1, 1] = "Entry Type";
                workSheet.Cells[1, 1] = "Regarding";



                if (!File.Exists(testPath))
                {
                    try
                    {
                        workBook.SaveAs(testPath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlShared, misValue, misValue, misValue, misValue, misValue);
                        workBook.Close(true, misValue, misValue);
                        excel.Quit();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("The Excel file: ImpNotes.xls could not be created because of the following error: \r\r", "Could Not Create Notes File In DOCS Folder", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }            
            catch 
            {

            }

            

            OleDbConnection connection = new OleDbConnection();
            try
            {
                //string excelTableSource = "Provider = Microsoft.Jet.OLEDB.4.0; Data Source = " + testPath + "; Extended Properties = 'Excel 8.0; HDR = YES'";

                //OleDbCommand command = new OleDbCommand();
                //string sql = null;

                //connection = new OleDbConnection(excelTableSource);

                //connection.Open();
                //command.Connection = connection;
                ////UPDATE table SET column = rowvalue WHERE columnToChange = value
                //sql = "UPDATE [Brandon$] SET " + columnName + " = @columnName WHERE " + locationColumn + " = @locationColumn;";

                //command.Parameters.AddWithValue("@columnName", value);
                //command.Parameters.AddWithValue("@locationColumn", erid);

                //command.CommandText = sql;
                //command.ExecuteNonQuery();
                //connection.Close();
                //sql = null;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Something prevented the data from pushing to the Excel database.\r\r" + ex, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
