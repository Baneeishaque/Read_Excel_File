using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Read_Excel_File
{
    class Program
    {
        static void Main(string[] args)
        {
            string Path = System.Windows.Forms.Application.StartupPath + "\\sample.xlsx";

            // initialize the Excel Application class
            ApplicationClass app = new ApplicationClass();

            // create the workbook object by opening the excel file.
            Workbook workBook = app.Workbooks.Open(Path, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

            // get the active worksheet using sheet name or active sheet
            Worksheet workSheet = (Worksheet)workBook.ActiveSheet;

            // This row,column index should be changed as per your need, i.e. which cell in the excel you are interesting to read.

            int row_index = 2;
            int first_name_column_index = 1;
            int last_name_column_index = 2;

            try
            {
                while (((Range)workSheet.Cells[row_index, first_name_column_index]).Value2 != null)
                {
                    string firstName =
                      ((Range)workSheet.Cells[row_index, first_name_column_index]).Value2.ToString();
                    string lastName =
                      ((Range)workSheet.Cells[row_index, last_name_column_index]).Value2.ToString();
                    Console.WriteLine("Name : {0} {1} ", firstName, lastName);
                    row_index++;
                }
                workBook.Close();
                app.Quit();
            }
            catch (Exception ex)
            {
                workBook.Close();
                app.Quit();
                Console.WriteLine(ex.Message);
            }

            Console.ReadKey();
        }
    }
}
