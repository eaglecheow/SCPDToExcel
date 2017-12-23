using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace SCPDRecorder.Services
{
    class ExcelExportService
    {
        public string Title { get; set; }

        public void StringListToExcel(List<string>stringList)
        {
            Console.WriteLine("Exporting file to Microsoft Excel...");

            var excel = new Application();
            excel.Visible = true;
            var workbook = excel.Workbooks.Add(Type.Missing);
            var sheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.ActiveSheet;
            sheet.Name = "s-CPD Record";
            sheet.Cells[1, 1] = "s-CPD Record List";

            sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, 2]].Merge();

            sheet.Cells[1, 1].Font.Bold = true;

            int startIndex = 2;

            foreach (string stringItem in stringList)
            {
                sheet.Cells[startIndex, 1] = stringItem;
                startIndex++;
            }

            sheet.Columns.AutoFit();
        }
    }
}
