using System;
using Aspose.Cells;

class ConvertWorkbook
{
    static void Main()
    {
        // Path to the source XLSX workbook
        string sourcePath = "input.xlsx";

        // Load the workbook (workbook-load rule)
        Workbook workbook = new Workbook(sourcePath);

        // Save the workbook as OpenDocument Spreadsheet (ODS) (workbook-save rule)
        workbook.Save("output.ods", SaveFormat.Ods);

        // Save the workbook as StarOffice Calc Spreadsheet (SXC) (workbook-save rule)
        workbook.Save("output.sxc", SaveFormat.Sxc);

        // Save the workbook as OpenDocument Flat XML Spreadsheet (FODS) (workbook-save rule)
        workbook.Save("output.fods", SaveFormat.Fods);
    }
}