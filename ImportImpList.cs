using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Microsoft.Office.Interop.Excel;

namespace FCHelper_v001
{
    class ImportImpList
    {
        public void ImportImpListMethod(string fileName)
        {
            System.Data.DataTable datatable = new System.Data.DataTable();
            //datatable.Clear();


            Workbook xlworkbook = new Workbook();
            Worksheet xlWorksheet = new Worksheet();
            Range rangeRowsRange = xlWorksheet.UsedRange.Rows;
            Range rangeColumnsRange = xlWorksheet.UsedRange.Columns;
            /*
            int

           
            
            for (var i = 1; i <= rangeColumns; i++)
            {
                var column = xlWorksheet.Cell(1, i);
                datatable.Columns.Add(column.Value.ToString());
            }

            var firstHeadRow = 0;
            foreach (var item in range.Rows())
            {
                if (firstHeadRow != 0)
                {
                    var array = new object[col];
                    for (var y = 1; y <= col; y++)
                    {
                        array[y - 1] = item.Cell(y).Value;
                    }

                    datatable.Rows.Add(array);
                }
                firstHeadRow++;
            }*/
        }
    }
}
