using System;
using System.Globalization;
using Aspose.Cells;

namespace LocalizationDemo
{
    // Custom globalization settings for Russian language
    public class RussianGlobalizationSettings : GlobalizationSettings
    {
        // Localize boolean values
        public override string GetBooleanValueString(bool value)
        {
            return value ? "ИСТИНА" : "ЛОЖЬ";
        }

        // Localize error values
        public override string GetErrorValueString(string err)
        {
            switch (err)
            {
                case "#NAME?":   return "#ИМЯ?";
                case "#DIV/0!":  return "#ДЕЛ/0!";
                case "#REF!":    return "#ССЫЛКА!";
                case "#VALUE!":  return "#ЗНАЧ!";
                case "#N/A":     return "#Н/Д";
                case "#NUM!":    return "#ЧИСЛО!";
                case "#NULL!":   return "#ПУСТО!";
                default:         return base.GetErrorValueString(err);
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file
            string inputPath = "input.xlsx";

            // Load the workbook with Russian culture settings
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            loadOptions.CultureInfo = new CultureInfo("ru-RU"); // ensures numeric/date parsing uses Russian locale
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Apply the custom Russian globalization settings
            workbook.Settings.GlobalizationSettings = new RussianGlobalizationSettings();

            // Example: write some test data
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Boolean values
            cells[0, 0].PutValue(true);
            cells[0, 1].PutValue(false);

            // Error values
            string[] errors = new string[] { "#NAME?", "#DIV/0!", "#REF!", "#VALUE!", "#N/A", "#NUM!", "#NULL!" };
            for (int i = 0; i < errors.Length; i++)
            {
                cells[0, i + 2].PutValue(errors[i]);
            }

            // Display localized strings in console
            Console.WriteLine("Localized cell values:");
            for (int col = 0; col < 9; col++)
            {
                Console.WriteLine($"Cell[0,{col}]: {cells[0, col].StringValue}");
            }

            // Save the localized workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}