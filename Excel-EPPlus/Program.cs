using OfficeOpenXml;
using System;
using System.IO;

namespace Excel.EPPlus
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage excel = new ExcelPackage();
            var workSheet = excel.Workbook.Worksheets.Add("Sheet1");

            for (int x = 1; x < 100; x++)
                for (int y = 1; y < 10; y++)
                {
                    workSheet.Cells[x, y].Value = Guid.NewGuid().ToString();
                }

            excel.SaveAs(new FileInfo("newWorkbook.xlsx"));
        }
    }
}