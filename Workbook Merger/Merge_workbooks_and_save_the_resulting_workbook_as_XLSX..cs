using System;
using Aspose.Cells;

namespace WorkbookMergeDemo
{
    class Program
    {
        static void Main()
        {
            // Paths for temporary workbooks
            string firstPath = "FirstWorkbook.xlsx";
            string secondPath = "SecondWorkbook.xlsx";
            string mergedPath = "MergedWorkbook.xlsx";

            // -------------------------------------------------
            // Create first workbook and add sample data
            // -------------------------------------------------
            Workbook firstWorkbook = new Workbook();                     // create
            Worksheet firstSheet = firstWorkbook.Worksheets[0];
            firstSheet.Name = "FirstSheet";
            firstSheet.Cells["A1"].PutValue("Data from first workbook");
            firstWorkbook.Save(firstPath, SaveFormat.Xlsx);            // save

            // -------------------------------------------------
            // Create second workbook and add sample data
            // -------------------------------------------------
            Workbook secondWorkbook = new Workbook();                    // create
            Worksheet secondSheet = secondWorkbook.Worksheets[0];
            secondSheet.Name = "SecondSheet";
            secondSheet.Cells["A1"].PutValue("Data from second workbook");
            secondWorkbook.Save(secondPath, SaveFormat.Xlsx);          // save

            // -------------------------------------------------
            // Load the workbooks to be merged
            // -------------------------------------------------
            Workbook destWorkbook = new Workbook(firstPath);            // load first as destination
            Workbook srcWorkbook = new Workbook(secondPath);            // load second as source

            // -------------------------------------------------
            // Merge the source workbook into the destination workbook
            // -------------------------------------------------
            destWorkbook.Combine(srcWorkbook);                          // combine

            // -------------------------------------------------
            // Save the merged workbook as XLSX
            // -------------------------------------------------
            destWorkbook.Save(mergedPath, SaveFormat.Xlsx);            // save merged result

            Console.WriteLine($"Workbooks merged successfully. Output file: {mergedPath}");
        }
    }
}