using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace SCPDRecorder
{
    class Program
    {
        static List<string> libraryNumberList = new List<string>();

        static void Main(string[] args)
        {
            RecordNumber();
            ExportExcel();
        }

        static void ExportExcel()
        {
            var excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = true;
            var workbook = excel.Workbooks.Add(Type.Missing);
            var sheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.ActiveSheet;
            sheet.Name = "s-CPD Record";

            int startIndex = 1;
            foreach (string libraryNumber in libraryNumberList)
            {
                sheet.Cells[startIndex, 1] = libraryNumber;
                startIndex++;
            }
        }

        static void RecordNumber()
        {
            bool exitStatus = false;
            while (!exitStatus)
            {
                Console.Write("Please enter library number : ");
                string libraryNumber = Console.ReadLine();
                if (!libraryNumber.Contains('q') && !libraryNumber.Contains('Q'))
                {
                    if (!libraryNumberList.Contains(libraryNumber))
                    {
                        libraryNumberList.Add(libraryNumber);
                        Console.WriteLine("Library number registered");
                    }
                    else
                    {
                        Console.WriteLine("Duplicate library number found. Current register ignored");
                    }
                }
                else
                {
                    exitStatus = true;
                }
            }
        }
    }
}
