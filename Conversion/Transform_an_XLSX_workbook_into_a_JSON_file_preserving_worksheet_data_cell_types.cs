using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook (workbook-load rule)
        Workbook workbook = new Workbook("input.xlsx");

        // Configure JSON save options to preserve structure and cell details
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            // Export the workbook as a JSON object even if it contains a single sheet
            AlwaysExportAsJsonObject = true,
            // Preserve parent‑child hierarchy (nested structure)
            ExportNestedStructure = true,
            // Include empty cells in the output (null values)
            ExportEmptyCells = true,
            // Export cell values as their native types (not forced to string)
            ExportAsString = false,
            // Skip empty rows is disabled to keep the original layout
            SkipEmptyRows = false
        };

        // Save the workbook as a JSON file (workbook-save rule)
        workbook.Save("output.json", jsonOptions);
    }
}