using Aspose.Cells;
using System;

class Program
{
    static void Main()
    {
        // Create LoadOptions instance
        LoadOptions loadOptions = new LoadOptions();

        // Configure the LoadFilter to load only cell values (no formulas, formatting, etc.)
        loadOptions.LoadFilter = new LoadFilter(LoadDataFilterOptions.CellValue);

        // Load the XLSX workbook using the specified LoadOptions
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Example: read a cell value to verify that data was loaded
        Worksheet sheet = workbook.Worksheets[0];
        Console.WriteLine("A1 value: " + sheet.Cells["A1"].StringValue);

        // Save the workbook (optional, here we save it unchanged)
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}