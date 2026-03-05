using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsExamples
{
    public class Program
    {
        public static void Main()
        {
            XlsxToXpsConversion.Run();
        }
    }

    public class XlsxToXpsConversion
    {
        public static void Run()
        {
            // Path to the source XLSX file
            string sourcePath = "input.xlsx";

            // Path for the output XPS file
            string destPath = "output.xps";

            // Convert the Excel file to XPS using Aspose.Cells ConversionUtility
            ConversionUtility.Convert(sourcePath, destPath);

            Console.WriteLine("Conversion from XLSX to XPS completed successfully.");
        }
    }
}