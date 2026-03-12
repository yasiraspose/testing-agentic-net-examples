using System;
using Aspose.Cells;

namespace AsposeCellsConversionDemo
{
    class Program
    {
        static void Main()
        {
            // Define input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            try
            {
                // Load the Excel workbook (workbook-load rule)
                Workbook workbook = new Workbook(inputPath);

                // Save the workbook as PDF (workbook-save rule with SaveFormat.Pdf)
                workbook.Save(outputPath, SaveFormat.Pdf);

                Console.WriteLine("Conversion completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
            }
        }
    }
}