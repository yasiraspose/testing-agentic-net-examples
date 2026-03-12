using System;
using Aspose.Cells;

namespace Example
{
    class Program
    {
        static void Main()
        {
            // Paths to the source and destination files
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Create LoadOptions and assign a custom filter that loads only visible sheets
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LoadFilter = new VisibleSheetsLoadFilter();

            // Load the workbook using the custom LoadOptions
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Verify which worksheets were loaded
            Console.WriteLine("Loaded worksheets:");
            foreach (Worksheet ws in workbook.Worksheets)
            {
                Console.WriteLine($"- {ws.Name} (Visible = {ws.IsVisible})");
            }

            // Save the workbook; only the visible sheets are present
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
    }

    // Custom LoadFilter that loads data only for visible worksheets
    class VisibleSheetsLoadFilter : LoadFilter
    {
        public override void StartSheet(Worksheet sheet)
        {
            // Load the sheet only if it is visible
            if (sheet.IsVisible)
            {
                base.StartSheet(sheet);
            }
        }
    }
}