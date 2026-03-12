using System;
using Aspose.Cells;               // Core Aspose.Cells namespace
using Aspose.Cells.Rendering;    // Contains XpsSaveOptions (inherits from PaginatedSaveOptions)

class XlsxToXpsConverter
{
    static void Main()
    {
        // Path to the source XLSX file
        string sourcePath = "input.xlsx";

        // Path where the XPS file will be saved
        string destPath = "output.xps";

        // Load the workbook (uses the workbook-load rule)
        Workbook workbook = new Workbook(sourcePath);

        // Create XPS save options – free‑form code (no specific rule exists)
        XpsSaveOptions saveOptions = new XpsSaveOptions
        {
            // Preserve layout by rendering each sheet on a separate page
            OnePagePerSheet = false,

            // Optional: ensure all columns fit on one page per sheet if desired
            // AllColumnsInOnePagePerSheet = true,

            // Use a default font to improve Unicode rendering consistency
            DefaultFont = "Arial",

            // Keep errors visible to aid debugging (set to false in production)
            IgnoreError = false
        };

        // Save the workbook as XPS using the workbook-save rule
        workbook.Save(destPath, saveOptions);

        Console.WriteLine($"Conversion completed: '{sourcePath}' → '{destPath}'");
    }
}