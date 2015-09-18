using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelEngine.NPOI
{
    internal class XlsxWriter : BaseWriter
    {
        public override void CreateWorkBook(string destinationPath)
        {
            try
            {
                var workbook = new XSSFWorkbook();
                var sheet1 = (XSSFSheet)workbook.CreateSheet("Sheet1");

                sheet1.CreateRow(0).CreateCell(0).SetCellValue("This is a Sample  xlsx Sheet");
                //var x = 1;
                //for (var i = 1; i <= 10000; i++)
                //{
                //    var row = sheet1.CreateRow(i);
                //    for (var j = 0; j < 50; j++)
                //    {
                //        row.CreateCell(j).SetCellValue("cell" + x++);
                //    }
                //}
                var s = sheet1 as ISheet;

                CreateRows(20000, 40, ref s);

                using (var f = File.Create(destinationPath))
                {
                    workbook.Write(f);
                }
            }
            catch (Exception e)
            {
                throw new Exception("Cannot create Excell SpreadSheet (.xlsx)", e);
            }
        }
    }
}
