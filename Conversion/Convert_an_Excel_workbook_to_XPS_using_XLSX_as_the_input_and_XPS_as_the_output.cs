using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source XLSX file
        string sourcePath = "input.xlsx";

        // Desired path for the XPS output
        string destPath = "output.xps";

        // Load the Excel workbook (uses workbook-load rule)
        Workbook workbook = new Workbook(sourcePath);

        // Save the workbook as XPS (uses workbook-save rule with SaveFormat)
        workbook.Save(destPath, SaveFormat.Xps);
    }
}