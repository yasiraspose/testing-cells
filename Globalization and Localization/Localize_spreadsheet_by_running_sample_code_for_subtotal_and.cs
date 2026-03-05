using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Load an existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Apply custom globalization settings for subtotal and grand total labels
        workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

        // Define the range to which the subtotal will be applied (e.g., A1:B5)
        CellArea area = CellArea.CreateCellArea(0, 0, 4, 1); // rows 0‑4, columns 0‑1

        // Apply subtotal: group by column 0, sum column 1, show subtotals and grand total
        cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 0 }, true, false, true);

        // Example: retrieve and display the localized grand total label for Sum
        string grandTotalLabel = ((CustomGlobalizationSettings)workbook.Settings.GlobalizationSettings)
                                 .GetGrandTotalName(ConsolidationFunction.Sum);
        Console.WriteLine("Localized Grand Total label: " + grandTotalLabel);

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }

    // Custom globalization settings overriding label texts
    class CustomGlobalizationSettings : GlobalizationSettings
    {
        // Localize the grand total label based on the consolidation function
        public override string GetGrandTotalName(ConsolidationFunction functionType)
        {
            return functionType == ConsolidationFunction.Sum ? "Total Général" : base.GetGrandTotalName(functionType);
        }

        // Localize subtotal labels (obsolete method retained for compatibility)
        public override string GetSubTotalName(PivotFieldSubtotalType subTotalType)
        {
            switch (subTotalType)
            {
                case PivotFieldSubtotalType.Sum:
                    return "Sous‑total Somme";
                case PivotFieldSubtotalType.Count:
                    return "Sous‑total Compte";
                case PivotFieldSubtotalType.Average:
                    return "Sous‑total Moyenne";
                default:
                    return base.GetSubTotalName(subTotalType);
            }
        }
    }
}