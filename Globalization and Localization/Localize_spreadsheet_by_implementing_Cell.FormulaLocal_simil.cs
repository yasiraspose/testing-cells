using System;
using Aspose.Cells;

namespace AsposeCellsFormulaLocalDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing XLSX workbook (replace with your actual file path)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Set the workbook's locale to German (de-DE) for demonstration
            workbook.Settings.Region = CountryCode.Germany;

            // Access the first worksheet and cell A1
            Worksheet worksheet = workbook.Worksheets[0];
            Cell cell = worksheet.Cells["A1"];

            // Example 1: Read the formula in standard (English) format
            Console.WriteLine("Standard Formula: " + cell.Formula);

            // Example 2: Read the formula in the locale‑specific format using FormulaLocal
            Console.WriteLine("Localized Formula (FormulaLocal): " + cell.FormulaLocal);

            // Example 3: Set a formula using the localized (German) syntax
            // German function name for SUM is SUMME
            cell.FormulaLocal = "=SUMME(B1:C1)";

            // Verify that the standard (English) formula has been translated automatically
            Console.WriteLine("\nAfter setting FormulaLocal:");
            Console.WriteLine("Standard Formula: " + cell.Formula);
            Console.WriteLine("Localized Formula: " + cell.FormulaLocal);

            // Recalculate the workbook so the formula result is updated
            workbook.CalculateFormula();

            // Display the calculated value of the cell
            Console.WriteLine("\nCalculated Value of A1: " + cell.Value);

            // Save the modified workbook (replace with your desired output path)
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"\nWorkbook saved to '{outputPath}'.");
        }
    }
}