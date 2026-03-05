using System;
using Aspose.Cells;

class RemoveThreadedComments
{
    static void Main()
    {
        // Load the workbook from an existing XLSX file
        Workbook workbook = new Workbook("input.xlsx");

        // Clear all comments (including threaded comments) from each worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.ClearComments();
        }

        // Save the workbook after removing the comments
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}