using System;
using Aspose.Cells;

namespace ExcelToHtmlConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            // Example usage:
            // Convert "input.xlsx" to "output.html"
            ConvertExcelToHtml("input.xlsx", "output.html");
        }

        /// <summary>
        /// Converts an Excel workbook to an HTML file while preserving formatting and content.
        /// </summary>
        /// <param name="sourcePath">Path to the source XLSX file.</param>
        /// <param name="outputPath">Path where the HTML file will be saved.</param>
        static void ConvertExcelToHtml(string sourcePath, string outputPath)
        {
            // Load the workbook from the specified file.
            // Rule: workbook-load
            Workbook workbook = new Workbook(sourcePath);

            // Configure HTML save options to retain formatting, images, and workbook properties.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                // Export workbook and worksheet properties for full fidelity.
                ExportWorkbookProperties = true,
                ExportWorksheetProperties = true,

                // Include grid lines to match the Excel view.
                ExportGridLines = true,

                // Embed images directly in the HTML as Base64 strings.
                ExportImagesAsBase64 = true,

                // Use HTML5 for modern standards compliance.
                HtmlVersion = HtmlVersion.Html5
            };

            // Save the workbook as HTML using the configured options.
            // Rule: workbook-save
            workbook.Save(outputPath, saveOptions);
        }
    }
}