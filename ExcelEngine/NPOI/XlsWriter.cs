using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NPOI.HSSF.Model;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace ExcelEngine.NPOI
{
    internal class XlsWriter : BaseWriter
    {
        public override void CreateWorkBook(string destinationPath)
        {
            try
            {
                var workbook = HSSFWorkbook.Create(InternalWorkbook.CreateWorkbook());

                // create sheet
                var sheet1 = (HSSFSheet) workbook.CreateSheet("Sheet1");

                sheet1.CreateRow(0).CreateCell(0).SetCellValue("This is a Sample xls Sheet");
                // 10000 rows, 50 columns
                //var x = 1;
                //for (var i = 0; i < 10000; i++)
                //{
                //    var r = sheet1.CreateRow(i);
                //    for (var j = 0; j < 50; j++)
                //    {
                //        r.CreateCell(j).SetCellValue("Cell " + x++);
                //    }
                //}
                var s = sheet1 as ISheet;
                CreateRows(20000, 40, ref s);

                using (var fs = File.Create(destinationPath))
                {
                    workbook.Write(fs);
                }
            }
            catch (Exception e)
            {
                throw new Exception("Cannot create Excell SpreadSheet (.xls)", e);
            }
        }
    }
}
    

