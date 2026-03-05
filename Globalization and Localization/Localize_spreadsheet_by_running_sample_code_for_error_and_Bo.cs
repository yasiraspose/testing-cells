using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing XLSX workbook
        string inputPath = "input.xlsx";
        Workbook wb = new Workbook(inputPath); // load rule

        // Prepare sample data: boolean values and error strings
        string[] errs = new string[] { "#NAME?", "#DIV/0!", "#REF!", "#VALUE!", "#N/A", "#NUM!", "#NULL!" };
        Cells cells = wb.Worksheets[0].Cells;

        cells[0, 0].PutValue(true);   // Boolean true
        cells[0, 1].PutValue(false);  // Boolean false

        for (int i = 0; i < errs.Length; i++)
        {
            cells[0, i + 2].PutValue(errs[i]); // Error strings
        }

        // Apply custom globalization settings for Boolean and Error localization
        wb.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

        // Output localized string values to console
        for (int i = 0; i < 9; i++)
        {
            Console.WriteLine($"Cell[0,{i}]: {cells[0, i].StringValue}");
        }

        // Save the modified workbook
        string outputPath = "output.xlsx";
        wb.Save(outputPath); // save rule
    }

    // Custom globalization settings overriding Boolean and Error string representations
    class CustomGlobalizationSettings : GlobalizationSettings
    {
        public override string GetBooleanValueString(bool bv)
        {
            return bv ? "ИСТИНА" : "ЛОЖЬ";
        }

        public override string GetErrorValueString(string err)
        {
            switch (err)
            {
                case "#NAME?": return "#ИМЯ?";
                case "#DIV/0!": return "#ДЕЛ/0!";
                case "#REF!": return "#ССЫЛКА!";
                case "#VALUE!": return "#ЗНАЧ!";
                case "#N/A": return "#Н/Д";
                case "#NUM!": return "#ЧИСЛО!";
                case "#NULL!": return "#ПУСТО!";
                default: return base.GetErrorValueString(err);
            }
        }
    }
}