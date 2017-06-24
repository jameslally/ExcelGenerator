using Excel.Npoi;
using Npoi.Core.HSSF.UserModel;
using Npoi.Core.HSSF.Util;
using Npoi.Core.SS.UserModel;
using Npoi.Core.SS.Util;
using Npoi.Core.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace ExcelNpoi
{
    public class Service
{

        public async Task<Stream> GenerateXlsx()
        {
            return await Generate(new XSSFWorkbook());
        }

        public async Task<Stream> GenerateXls()
        {
            return await Generate(new HSSFWorkbook());
        }

        public async Task<Stream> Generate(IWorkbook workbook)
    {

            var stream = new MemoryStreamNpoi();

            await Task.Run(() =>
            {
                ISheet sheet1 = workbook.CreateSheet("Sheet1");

                sheet1.AddMergedRegion(new CellRangeAddress(0, 0, 0, 10));
                var rowIndex = 0;
                IRow row = sheet1.CreateRow(rowIndex);
                row.Height = 30 * 80;
                row.CreateCell(0).SetCellValue("this is content, very long content, very long content, very long content, very long content");
                sheet1.AutoSizeColumn(0);
                rowIndex++;


                var sheet2 = workbook.CreateSheet("Sheet2");
                var style1 = workbook.CreateCellStyle();
                style1.FillForegroundColor = HSSFColor.Blue.Index2;
                style1.FillPattern = FillPattern.SolidForeground;

                var style2 = workbook.CreateCellStyle();
                style2.FillForegroundColor = HSSFColor.Yellow.Index2;
                style2.FillPattern = FillPattern.SolidForeground;

                var cell2 = sheet2.CreateRow(0).CreateCell(0);
                cell2.CellStyle = style1;
                cell2.SetCellValue(0);

                cell2 = sheet2.CreateRow(1).CreateCell(0);
                cell2.CellStyle = style2;
                cell2.SetCellValue(1);

                //Work around to a Java issue
                //https://stackoverflow.com/questions/22931582/memorystream-seems-be-closed-after-npoi-workbook-write#37398007
                stream.AllowClose = false;
                workbook.Write(stream);
                stream.AllowClose = true;
            });

            await stream.FlushAsync();

            if (stream != null)
                stream.Position = 0;
            return stream;
            
        }
    }
}

