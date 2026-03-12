using System;
using Aspose.Cells;

class ConvertXlsxToMhtml
{
    static void Main()
    {
        // Path to the source XLSX workbook
        string sourcePath = "input.xlsx";

        // Desired path for the resulting MHTML file
        string destinationPath = "output.mht";

        // Load the existing workbook (lifecycle rule: workbook-load)
        Workbook workbook = new Workbook(sourcePath);

        // Save the workbook as MHTML, preserving all data and formatting
        // (lifecycle rule: workbook-save with SaveFormat.MHtml)
        workbook.Save(destinationPath, SaveFormat.MHtml);

        Console.WriteLine("Workbook successfully converted to MHTML.");
    }
}