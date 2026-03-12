using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the existing XLSX file
        string inputPath = "example.xlsx";

        // Load the workbook using the provided lifecycle rule
        Workbook workbook = new Workbook(inputPath);

        // Access the first worksheet (optional demonstration)
        Worksheet sheet = workbook.Worksheets[0];
        Console.WriteLine($"Loaded worksheet: {sheet.Name}");
    }
}