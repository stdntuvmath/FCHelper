using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace FCHelper_v001
{
    class GetSpecificDataFromExcelDatabase
    {
        public void GetSpecificDataFromExcelDatabaseMethod(string ERID)
        {
            PrivateGetSpecificDataFromExcelDatabaseMethod(ERID);
        }

        private void PrivateGetSpecificDataFromExcelDatabaseMethod(string ERID)
        {
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();

            ExcelApp.Visible = false;

            string excelDBFilePath = @"C:\Users\14025\Documents\File Consultants\ImpList.xls";

            ExcelApp.Visible = false;

            Microsoft.Office.Interop.Excel.Workbook ExcelWorkbook = ExcelApp.Workbooks.Open(excelDBFilePath);
            ExcelApp.Visible = false;

            Microsoft.Office.Interop.Excel.Worksheet ExcelWorksheet = ExcelWorkbook.Sheets[1];


            Microsoft.Office.Interop.Excel.Range colRange = ExcelWorksheet.Columns["B:B"];//get the range object where you want to search from
            MessageBox.Show("ERID inside PrivateGetSpecificDataFromExcelDatabaseMethod: " + ERID);
            string searchString = ERID;
            ExcelApp.Visible = false;


            Microsoft.Office.Interop.Excel.Range resultRange = colRange.Find(

                What: searchString,

                LookIn: Microsoft.Office.Interop.Excel.XlFindLookIn.xlValues,

                LookAt: Microsoft.Office.Interop.Excel.XlLookAt.xlPart,

                SearchOrder: Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows,

                SearchDirection: Microsoft.Office.Interop.Excel.XlSearchDirection.xlNext

                );// search searchString in the range, if find result, return a range



            if (ERID == "")

            {

                MessageBox.Show("ERID " + searchString + " not provided.");

            }
            else
            {
                MessageBox.Show("ERID inside else statement"+ERID);
                string[] row = resultRange.EntireRow.Parse();
                foreach (string cell in row)
                {
                    MessageBox.Show(cell);
                }

                ExcelApp.Quit();
            }
        }
    }
}
