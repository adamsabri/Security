using System;
using System.Collections.Generic;
using System.IO;
using ExcelEngine.Exceptions;
using OfficeOpenXml;

namespace ExcelEngine.EPPlus
{
    internal class XlsxReader : IReader
    {
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
                throw new EpplusInitializeException("Cannot initialize Reader", e);
            }
        }

        public IList<string> ConvertToList()
        {
            try
            {
                var list = new List<string>();
                using (var package = new ExcelPackage(new FileInfo(FilePath)))
                {
                    // get the first worksheet in the workbook
                    var worksheet = package.Workbook.Worksheets[1];
                    var totalRows = worksheet.Dimension.End.Row;
                    var totalCols = worksheet.Dimension.End.Column;
                    for (var row = 1; row <= totalRows; row++)
                    {
                        var concat = string.Empty;
                        for (var col = 1; col <= totalCols; col++)
                            concat += worksheet.Cells[row, col].Value + " ";
                        list.Add(concat);
                    }
                }
                return list;
            }
            catch (Exception e)
            {
                throw new EpplusXlsxReaderException("Error while reading data from Sheets", e);
            }
        }
    }
}
