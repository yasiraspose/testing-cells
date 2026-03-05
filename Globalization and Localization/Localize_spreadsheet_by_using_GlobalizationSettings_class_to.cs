using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Pivot;

namespace AsposeCellsGlobalizationDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");

            // ------------------------------------------------------------
            // 1. Create globalization settings for subtotal (total) labels
            // ------------------------------------------------------------
            SettableGlobalizationSettings globalization = new SettableGlobalizationSettings();

            // Customize the total name for the SUM function (used in subtotals)
            globalization.SetTotalName(ConsolidationFunction.Sum, "Custom Sum Total");

            // ------------------------------------------------------------
            // 2. Create chart globalization settings for a pie chart
            // ------------------------------------------------------------
            SettableChartGlobalizationSettings chartGlobals = new SettableChartGlobalizationSettings();

            // Customize various chart texts
            chartGlobals.SetSeriesName("Custom Series");
            chartGlobals.SetChartTitleName("Custom Pie Chart Title");
            chartGlobals.SetLegendTotalName("Custom Total");
            chartGlobals.SetOtherName("Other Category");
            chartGlobals.SetLegendIncreaseName("Increase");
            chartGlobals.SetLegendDecreaseName("Decrease");

            // Assign the chart settings to the main globalization object
            globalization.ChartSettings = chartGlobals;

            // ------------------------------------------------------------
            // 3. Apply the globalization settings to the workbook
            // ------------------------------------------------------------
            workbook.Settings.GlobalizationSettings = globalization;

            // ------------------------------------------------------------
            // 4. (Optional) Verify that the workbook contains a pie chart.
            //    If needed, you can access the chart to force a refresh.
            // ------------------------------------------------------------
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Chart chart in sheet.Charts)
                {
                    if (chart.Type == ChartType.Pie || chart.Type == ChartType.PieExploded)
                    {
                        // Force the chart to re-evaluate its labels after globalization change
                        chart.NSeries[0].IsColorVaried = chart.NSeries[0].IsColorVaried;
                    }
                }
            }

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}