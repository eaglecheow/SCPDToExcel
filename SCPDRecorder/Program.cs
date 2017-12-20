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
            Console.WriteLine("===================");
            Console.WriteLine("s-CPD Recorder v1.0");
            Console.WriteLine("===================");
            Console.WriteLine();
            Console.WriteLine("Enter 'Q' to stop recording");
            Console.WriteLine();
            RecordNumber();
            ExportExcel();
            Console.WriteLine("Press any key to quit...");
            Console.ReadKey();
        }

        /// <summary>
        /// This function exports the list to an excel file
        /// </summary>
        static void ExportExcel()
        {
            Console.WriteLine("Exporting data to excel...");

            var excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = true;
            var workbook = excel.Workbooks.Add(Type.Missing);
            var sheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.ActiveSheet;
            sheet.Name = "s-CPD Record";
            sheet.Cells[1, 1] = "s-CPD Record List";

            sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, 2]].Merge();

            sheet.Cells[1, 1].Font.Bold = true;

            int startIndex = 2;
            foreach (string libraryNumber in libraryNumberList)
            {
                sheet.Cells[startIndex, 1] = libraryNumber;
                startIndex++;
            }

            sheet.Columns.AutoFit();
        }

        /// <summary>
        /// Records valid library numbers to list
        /// </summary>
        static void RecordNumber()
        {
            bool exitStatus = false;
            while (!exitStatus)
            {
                Console.Write("Please enter library number : ");
                string libraryNumber = Console.ReadLine();
                if (!libraryNumber.Contains('q') && !libraryNumber.Contains('Q'))
                {
                    if (string.IsNullOrEmpty(libraryNumber) || !libraryNumber.Any(char.IsDigit))
                    {
                        Console.WriteLine("Invalid data entered, please try again");
                    }
                    else if (!libraryNumberList.Contains(libraryNumber))
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

                Console.WriteLine();
            }
        }
    }
}
