using System;
using Aspose.Cells;

class Program
{
    // Custom LoadFilter that loads only visible worksheets
    class VisibleSheetsLoadFilter : LoadFilter
    {
        public override void StartSheet(Worksheet sheet)
        {
            // If the sheet is visible, load all its data
            if (sheet.IsVisible)
            {
                base.StartSheet(sheet);
            }
            else
            {
                // For hidden sheets load only the structure (no cell data)
                LoadDataFilterOptions = LoadDataFilterOptions.Structure;
            }
        }
    }

    static void Main()
    {
        string inputPath = "InputWorkbook.xlsx";
        string outputPath = "VisibleOnlyWorkbook.xlsx";

        // Create LoadOptions and assign the custom filter
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LoadFilter = new VisibleSheetsLoadFilter();

        // Load the workbook with the custom filter
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // Optional: display loaded worksheets information
        Console.WriteLine("Loaded worksheets:");
        foreach (Worksheet ws in workbook.Worksheets)
        {
            Console.WriteLine($"- {ws.Name} (Visible = {ws.IsVisible})");
        }

        // Save the workbook (only visible sheets contain data)
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}