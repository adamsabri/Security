using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace ExcelEngine.NPOI
{
    public abstract class BaseWriter : IWriter
    {
        public abstract void CreateWorkBook(string destinationPath);
        
        public void CreateRows(int rowsCount, int cellsCount, ref ISheet sheet, IList<string> cellsValues = null)
        {
            if (sheet == null) return;
            for (var i = 1; i <= rowsCount; i++)
            {
                var r = sheet.CreateRow(i);
                CreateCells(cellsCount, ref r, cellsValues);
            }
        }

        public void CreateCells(int cellsCount, ref IRow row, IList<string> cellsValues = null)
        {
            //var isEmptyCells = cellsValues == null || !cellsValues.Any() || cellsValues.Count != cellsCount;
            //if (isEmptyCells)
            //    for (var j = 0; j < cellsCount; j++)
            //        row.CreateCell(j).SetCellValue(string.Empty);
            //else
                for (var j = 0; j < cellsCount; j++)
                    row.CreateCell(j).SetCellValue("cell " + j);

        }
    }
}
