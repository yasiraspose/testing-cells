using System;
using Aspose.Cells;

namespace RemoveThreadedCommentsDemo
{
    class Program
    {
        static void Main()
        {
            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                // Determine the used range to limit iteration
                int maxRow = worksheet.Cells.MaxDataRow;
                int maxColumn = worksheet.Cells.MaxDataColumn;

                // Loop through each cell in the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxColumn; col++)
                    {
                        // Retrieve the threaded comments for the current cell
                        ThreadedCommentCollection threadedComments = worksheet.Comments.GetThreadedComments(row, col);

                        // If there are any threaded comments, clear them
                        if (threadedComments != null && threadedComments.Count > 0)
                        {
                            threadedComments.Clear();
                        }
                    }
                }
            }

            // Save the modified workbook (lifecycle rule: save)
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
    }
}