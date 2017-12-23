using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SCPDRecorder.Services;

namespace SCPDRecorder
{
    class Program
    {
        static List<string> libraryNumberList = new List<string>();

        static void Main(string[] args)
        {
            RecordNumber();
            ExcelExport();
            Console.WriteLine("Press any key to quit...");
            Console.ReadKey();
        }

        static void Menu()
        {
            var exitStatus = false;
            while (!exitStatus)
            {
                Console.WriteLine(@"Please select a function : 
1. Record Number
2. Export to Excel");
                Console.Write("Your option : ");
                var userInput = Console.ReadLine();
                exitStatus = true;
            }
        }

        static void ExcelExport()
        {
            ExcelExportService excelExport = new ExcelExportService();
            excelExport.StringListToExcel(libraryNumberList);
        }

        /// <summary>
        /// Records valid library numbers to list
        /// </summary>
        static void RecordNumber()
        {
            bool exitStatus = false;
            while (!exitStatus)
            {
                InputValidator inputValidator = new InputValidator();
                Console.Write("Please enter library number : ");
                string libraryNumber = Console.ReadLine();
                if (libraryNumber == "Q" || libraryNumber == "q")
                {
                    if (inputValidator.LibraryNumberValidator(libraryNumber))
                    {
                        if (!libraryNumberList.Contains(libraryNumber))
                        {
                            libraryNumberList.Add(libraryNumber);
                        }
                        else
                        {
                            Console.WriteLine("User already registered");
                        }
                    }
                    else
                    {
                        Console.WriteLine("Invalid input. Please try again.");
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
