using System;
using Aspose.Cells;

namespace AsposeCellsFormulaLocalDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source XLSX workbook
            string inputPath = "input.xlsx";

            // Load the workbook (uses the standard load lifecycle)
            Workbook workbook = new Workbook(inputPath);

            // Set the workbook region to German to obtain German localized formulas
            workbook.Settings.Region = CountryCode.Germany;

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Iterate through all used cells and display formula information
            foreach (Cell cell in cells)
            {
                // Process only cells that contain a formula
                if (!string.IsNullOrEmpty(cell.Formula))
                {
                    // Cell name (e.g., A1)
                    string address = cell.Name;

                    // Standard (English) formula
                    string standardFormula = cell.Formula;

                    // Localized formula obtained via the FormulaLocal property
                    string localizedFormula = cell.FormulaLocal;

                    // Localized formula obtained via GetFormula (isLocal = true)
                    string localizedViaGet = cell.GetFormula(false, true);

                    // Output the information to the console
                    Console.WriteLine($"Cell {address}:");
                    Console.WriteLine($"  Standard Formula : {standardFormula}");
                    Console.WriteLine($"  FormulaLocal     : {localizedFormula}");
                    Console.WriteLine($"  GetFormula(true) : {localizedViaGet}");
                    Console.WriteLine();
                }
            }

            // Save the workbook (demonstrates the save lifecycle)
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
        }
    }
}