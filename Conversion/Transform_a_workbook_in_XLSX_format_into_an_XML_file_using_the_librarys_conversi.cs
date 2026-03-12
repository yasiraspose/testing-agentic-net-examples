using System;
using Aspose.Cells;

namespace AsposeCellsConversionDemo
{
    public class XlsxToXmlConverter
    {
        public static void Run()
        {
            // Path to the source XLSX file
            string sourcePath = "input.xlsx";

            // Desired path for the output XML file
            string outputPath = "output.xml";

            // Load the workbook from the XLSX file
            Workbook workbook = new Workbook(sourcePath);

            // Save the workbook as an XML file
            workbook.Save(outputPath, SaveFormat.Xml);

            Console.WriteLine($"Conversion completed: '{sourcePath}' -> '{outputPath}'");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            XlsxToXmlConverter.Run();
        }
    }
}