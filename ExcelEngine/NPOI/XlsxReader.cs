using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using ExcelEngine.Exceptions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelEngine.NPOI
{
    internal sealed class XlsxReader : IReader
    {
        private XSSFWorkbook _xssfworkbook;
        public string FilePath { get; private set; }

        public XlsxReader(string filePath)
        {
            try
            {
                if (!File.Exists(filePath))
                    throw new FileNotFoundException();
                FilePath = filePath;
            }
            catch (Exception e)
            {
                throw new NpoiInitializeException("Cannot initialize Xlsx Reader", e);
            }
        }

        public IList<string> ConvertToList()
        {
            try
            {
                OpenWorkBook();

                // get sheet
                //sh = (XSSFSheet)wb.GetSheet(name);
                var sh = (XSSFSheet) _xssfworkbook.GetSheetAt(0);

                var list = new List<string>();
                var i = 0;
                while (sh.GetRow(i) != null)
                {
                    var concat = string.Empty;
                    // write row value
                    for (var j = 0; j < sh.GetRow(i).Cells.Count; j++)
                    {
                        var cell = sh.GetRow(i).GetCell(j);

                        if (cell == null) continue;
                        // TODO: you can add more cell types capability, e. g. formula
                        switch (cell.CellType)
                        {
                            case CellType.Numeric:
                                concat = sh.GetRow(i).GetCell(j).NumericCellValue.ToString(CultureInfo.InvariantCulture);
                                break;
                            case CellType.String:
                                concat = sh.GetRow(i).GetCell(j).StringCellValue;
                                break;
                        }
                    }
                    list.Add(concat);
                    i++;
                }
                return list;
            }
            catch (NpoiInitializeException e)
            {
                throw new NpoiInitializeException(e.Message, e);
            }
            catch (Exception e)
            {
                throw new NpoiXlsxReaderException("Error while reading data from Sheets", e);
            }
        }

        private void OpenWorkBook()
        {
            try
            {
                using (var file = new FileStream(FilePath, FileMode.Open, FileAccess.Read))
                    _xssfworkbook = new XSSFWorkbook(file);
            }
            catch (Exception e)
            {
                throw new NpoiInitializeException("Cannot initialize Xlsx Workbook", e);
            }
        }

    }
}
