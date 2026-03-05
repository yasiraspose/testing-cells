using System;
using Aspose.Cells;

class WorkbookProtectionDemo
{
    static void Main()
    {
        // Load the existing workbook
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Protect the workbook (structure) with a password
        string password = "myPassword123";
        workbook.Protect(ProtectionType.Structure, password);

        // Save the protected workbook
        string protectedPath = "protected.xlsx";
        workbook.Save(protectedPath, SaveFormat.Xlsx);

        // Load the protected workbook
        Workbook protectedWorkbook = new Workbook(protectedPath);

        // Unprotect the workbook using the same password
        protectedWorkbook.Unprotect(password);

        // Save the unprotected workbook
        string unprotectedPath = "unprotected.xlsx";
        protectedWorkbook.Save(unprotectedPath, SaveFormat.Xlsx);
    }
}