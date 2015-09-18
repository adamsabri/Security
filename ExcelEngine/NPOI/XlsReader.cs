using System;
using System.Collections.Generic;
using System.IO;
using ExcelEngine.Exceptions;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace ExcelEngine.NPOI
{
    internal sealed class XlsReader : IReader
    {
        private HSSFWorkbook _hssfworkbook;
        public string FilePath { get; private set; }

        public XlsReader(string filePath)
        {
           try
            {
                if (!File.Exists(filePath))
                    throw new FileNotFoundException();
                FilePath = filePath;
            }
            catch (Exception e)
            {
                throw new NpoiInitializeException("Cannot initialize Xls Reader", e);
            }
        }

        public IList<string> ConvertToList()
        {
            try
            {
                OpenWorkBook();

                var sheet = _hssfworkbook.GetSheetAt(0);
                var rows = sheet.GetRowEnumerator();

                var list = new List<string>();
                while (rows.MoveNext())
                {
                    IRow row = (HSSFRow)rows.Current;
                    var concat = string.Empty;
                    for (var i = 0; i < row.LastCellNum; i++)
                    {
                        var cell = row.GetCell(i);
                        concat += (cell == null ? null : cell.ToString()) + " ";
                    }
                    list.Add(concat);
                }
                return list;
            }
            catch (NpoiInitializeException e)
            {
                throw new NpoiInitializeException(e.Message, e);
            }
            catch (Exception e)
            {
                throw new NpoiXlsReaderException("Error while reading data from Sheets", e);
            }
        }

        private void OpenWorkBook()
        {
            try
            {
                //read the template via FileStream, it is suggested to use FileAccess.Read to prevent file lock.
                using (var file = new FileStream(FilePath, FileMode.Open, FileAccess.Read))
                    _hssfworkbook = new HSSFWorkbook(file, false);
            }
            catch (Exception e)
            {
                throw new NpoiInitializeException("Cannot initialize Xls Workbook", e);
            }
        }
    }
}
