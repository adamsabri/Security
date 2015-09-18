using System.Collections.Generic;
using System.Linq;
using NPOI.SS.UserModel;

namespace ExcelEngine.NPOI
{
    public interface IWriter
    {
        void CreateWorkBook(string destinationPath);
        void CreateRows(int rowsCount, int cellsCount, ref ISheet sheet, IList<string> cellsValues = null);
        void CreateCells(int cellsCount, ref IRow row, IList<string> cellsValues = null);
    }
}
