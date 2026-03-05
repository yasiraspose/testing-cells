using System;
using Aspose.Cells;

namespace FormulaLocalLocalizationDemo
{
    class Program
    {
        static void Main()
        {
            // 1. Load an existing workbook (replace with your actual file path)
            string inputPath = "InputWorkbook.xlsx";
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // 2. Access a cell that already contains a formula
            Cell cellA1 = sheet.Cells["A1"];
            Console.WriteLine("=== Existing Formula ===");
            Console.WriteLine($"Standard Formula : {cellA1.Formula}");
            Console.WriteLine($"Localized Formula: {cellA1.FormulaLocal}");

            // 3. Change workbook region to German and observe FormulaLocal change
            workbook.Settings.Region = CountryCode.Germany;
            // Re‑read the same cell after region change
            Console.WriteLine("\n=== After Setting Region to Germany ===");
            Console.WriteLine($"Standard Formula : {cellA1.Formula}");
            Console.WriteLine($"Localized Formula: {cellA1.FormulaLocal}");

            // 4. Set a formula using the localized (German) function name via FormulaLocal
            Cell cellB1 = sheet.Cells["B1"];
            cellB1.FormulaLocal = "=SUMME(C1:C5)"; // German name for SUM
            // Populate the argument range
            for (int i = 0; i < 5; i++)
                sheet.Cells[$"C{i + 1}"].PutValue(i + 1);
            // Verify that the standard Formula reflects the English name
            Console.WriteLine("\n=== After Setting FormulaLocal (German) ===");
            Console.WriteLine($"Standard Formula : {cellB1.Formula}");
            Console.WriteLine($"Localized Formula: {cellB1.FormulaLocal}");

            // 5. Use custom globalization settings to map a new local function name
            SettableGlobalizationSettings customSettings = new SettableGlobalizationSettings();
            customSettings.SetLocalFunctionName("AVERAGE", "MEDIE", true); // French for AVERAGE
            workbook.Settings.GlobalizationSettings = customSettings;

            Cell cellC1 = sheet.Cells["C1"];
            cellC1.FormulaLocal = "=MEDIE(D1:D4)"; // Use the custom local name
            // Populate the argument range
            sheet.Cells["D1"].PutValue(10);
            sheet.Cells["D2"].PutValue(20);
            sheet.Cells["D3"].PutValue(30);
            sheet.Cells["D4"].PutValue(40);
            Console.WriteLine("\n=== After Applying Custom Globalization Settings ===");
            Console.WriteLine($"Standard Formula : {cellC1.Formula}");
            Console.WriteLine($"Localized Formula: {cellC1.FormulaLocal}");

            // 6. Set a locale‑formatted formula directly using FormulaParseOptions
            FormulaParseOptions options = new FormulaParseOptions
            {
                LocaleDependent = true,
                R1C1Style = false
            };
            // French date format example
            sheet.Cells["E1"].SetFormula("=TEXT(AUJOURDHUI();\"[$-fr-FR]dddd, dd mmmm yyyy\")", options);
            Console.WriteLine("\n=== Formula Set with LocaleDependent Option ===");
            Console.WriteLine($"Standard Formula : {sheet.Cells["E1"].Formula}");
            Console.WriteLine($"Localized Formula: {sheet.Cells["E1"].FormulaLocal}");

            // 7. Calculate all formulas to ensure they work with the localized names
            workbook.CalculateFormula();

            // 8. Save the modified workbook
            string outputPath = "LocalizedDemoOutput.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"\nWorkbook saved to '{outputPath}'.");
        }
    }
}