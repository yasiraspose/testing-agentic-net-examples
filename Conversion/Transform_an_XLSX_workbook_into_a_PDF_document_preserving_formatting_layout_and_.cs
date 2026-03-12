using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class XlsxToPdfConverter
    {
        /// <summary>
        /// Converts an existing XLSX workbook to PDF while preserving formatting, layout and pagination.
        /// </summary>
        /// <param name="sourcePath">Full path to the source .xlsx file.</param>
        /// <param name="destPath">Full path where the resulting .pdf file will be saved.</param>
        public static void Convert(string sourcePath, string destPath)
        {
            // Load the workbook from the specified XLSX file.
            // Rule: workbook-load
            Workbook workbook = new Workbook(sourcePath);

            // Save the workbook as PDF. The Save method with SaveFormat.Pdf
            // preserves all formatting, page setup, and pagination.
            // Rule: workbook-save
            workbook.Save(destPath, SaveFormat.Pdf);

            Console.WriteLine($"Conversion completed: '{sourcePath}' → '{destPath}'");
        }

        // Example usage
        public static void Main()
        {
            string source = "input.xlsx";   // replace with your source file path
            string output = "output.pdf";   // replace with desired PDF output path

            Convert(source, output);
        }
    }
}