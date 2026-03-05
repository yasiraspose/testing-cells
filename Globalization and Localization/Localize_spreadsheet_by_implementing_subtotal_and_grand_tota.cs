using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Settings;

class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Create a SettableGlobalizationSettings instance to customize total labels
        SettableGlobalizationSettings globalization = new SettableGlobalizationSettings();

        // Localize the total and grand total names for the Sum function (example: Japanese)
        globalization.SetTotalName(ConsolidationFunction.Sum, "合計");          // "Total"
        globalization.SetGrandTotalName(ConsolidationFunction.Sum, "総計");   // "Grand Total"

        // Create a SettablePivotGlobalizationSettings instance to customize subtotal texts
        SettablePivotGlobalizationSettings pivotGlobalization = new SettablePivotGlobalizationSettings();

        // Localize various subtotal types
        pivotGlobalization.SetTextOfSubTotal(PivotFieldSubtotalType.Sum, "小計（合計）");       // "Subtotal (Sum)"
        pivotGlobalization.SetTextOfSubTotal(PivotFieldSubtotalType.Count, "小計（件数）");    // "Subtotal (Count)"
        pivotGlobalization.SetTextOfSubTotal(PivotFieldSubtotalType.Average, "小計（平均）"); // "Subtotal (Average)"

        // Attach the pivot globalization settings to the main globalization settings
        globalization.PivotSettings = pivotGlobalization;

        // Apply the customized globalization settings to the workbook
        workbook.Settings.GlobalizationSettings = globalization;

        // Define the range on which to apply the Subtotal operation (adjust as needed)
        // Here we assume data occupies rows 1‑6 (0‑5 index) and columns A‑B (0‑1 index)
        CellArea area = CellArea.CreateCellArea(0, 0, 5, 1);

        // Apply Subtotal: group by column 0 (A), calculate Sum on column 1 (B)
        // The last 'true' adds a grand total row
        cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 0 }, true, false, true);

        // Save the localized workbook
        workbook.Save("output_localized.xlsx");
    }
}