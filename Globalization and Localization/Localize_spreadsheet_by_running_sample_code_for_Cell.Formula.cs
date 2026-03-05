using System;
using System.IO;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsFormulaLocalDemo
{
    class Program
    {
        static void Main()
        {
            string inputPath = "SampleInput.xlsx";
            Workbook workbook;

            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                ws.Cells["B1"].PutValue(10);
                ws.Cells["C1"].PutValue(20);
                workbook.Save(inputPath);
            }

            // Set the workbook's default locale to German
            workbook.Settings.CultureInfo = new CultureInfo("de-DE");

            Worksheet worksheet = workbook.Worksheets[0];
            Cell cell = worksheet.Cells["A1"];

            // Set a formula using the standard (English) syntax
            cell.Formula = "=SUM(B1:C1)";

            Console.WriteLine("Standard Formula : " + cell.Formula);
            Console.WriteLine("Localized Formula: " + cell.FormulaLocal);

            // Set the formula using the German localized syntax
            cell.FormulaLocal = "=SUMME(B1:C1)";

            Console.WriteLine("\nAfter setting FormulaLocal:");
            Console.WriteLine("Standard Formula : " + cell.Formula);
            Console.WriteLine("Localized Formula: " + cell.FormulaLocal);

            // Demonstrate GetFormula with localization flag
            Console.WriteLine("\nUsing GetFormula:");
            Console.WriteLine("English formula   : " + cell.GetFormula(false, false));
            Console.WriteLine("Localized formula : " + cell.GetFormula(false, true));

            // Save the modified workbook
            string outputPath = "LocalizedOutput.xlsx";
            workbook.Save(outputPath);
        }
    }
}