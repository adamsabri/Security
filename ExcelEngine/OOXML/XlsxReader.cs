using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelEngine.Exceptions;

namespace ExcelEngine.OOXML
{
    internal class XlsxReader : IReader
    {
        public string FilePath { get; private set; }

        private readonly OooxmlAproach _aproach = OooxmlAproach.Sax;

        public XlsxReader(string filePath, OooxmlAproach aproach = OooxmlAproach.Sax)
        {
            try
            {
                if (!File.Exists(filePath))
                    throw new FileNotFoundException();
                FilePath = filePath;
                _aproach = aproach;
            }
            catch (Exception e)
            {
                throw new OoxmlInitializeException("Cannot initialize Reader", e);
            }
        }

        public IList<string> ConvertToList()
        {
            switch (_aproach)
            {
                case OooxmlAproach.Dom:
                    return ReadOnDom();
                default:
                    return ReadOnSax();
            }
        }

        /// <summary>
        /// The DOM approach.
        /// The method above will definitely work for reading cell values within a worksheet.
        /// However, if the worksheet is quite large then the program's memory footprint will
        /// also be quite large. In fact, you are left at the mercy of the garbage collector,
        /// which may result in the program throwing an Out of Memory exception. 
        /// Note that the code below works only for cells that contain numeric values.
        /// </summary>
        private IList<string> ReadOnDom()
        {
            try
            {
                using (var spreadsheetDocument = SpreadsheetDocument.Open(FilePath, false))
                {
                    var workbookPart = spreadsheetDocument.WorkbookPart;
                    var worksheetPart = workbookPart.WorksheetParts.First();
                    var sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                    var list = new List<string>();
                    var concat = string.Empty;
                    foreach (var r in sheetData.Elements<Row>())
                    {
                        concat = string.Empty;
                        foreach (var c in r.Elements<Cell>())
                        {
                           concat += ReadExcelCell(c, ref workbookPart) + " ";
                        }
                    }
                    list.Add(concat);
                    return list;
                }
            }
            catch (Exception e)
            {
                throw new OoxmlXlsxDomReaderException("Error while reading data from Sheets", e);
            }
        }

        /// <summary>
        /// The SAX approach.
        /// The method above takes advantage of the OpenXmlReader class,
        /// which allows you to read an entire part without loading the part in memory.
        /// </summary>
        private IList<string> ReadOnSax()
        {
            try
            {
                using (var spreadsheetDocument = SpreadsheetDocument.Open(FilePath, false))
                {
                    var workbookPart = spreadsheetDocument.WorkbookPart;
                    var worksheetPart = workbookPart.WorksheetParts.First();

                    var reader = OpenXmlReader.Create(worksheetPart);
                    var i = 0;
                    var concat = string.Empty;
                    var list = new List<string>();
                    while (reader.Read())
                    {
                        if (reader.ElementType == typeof(Cell))
                        {
                            var c = (Cell)reader.LoadCurrentElement();
                            concat += ReadExcelCell(c, ref workbookPart) + " ";
                        }
                        else if (reader.ElementType == typeof(Row))
                        {
                            if (i > 1 && !string.IsNullOrEmpty(concat))
                                list.Add(concat);
                            concat = string.Empty;
                            i++;
                        }
                    }
                    return list;
                }
            }
            catch (Exception e)
            {
                throw new OoxmlXlsxSaxReaderException("Error while reading data from Sheets", e);
            }
        }

        private static string ReadExcelCell(Cell cell, ref WorkbookPart workbookPart)
        {
            var cellValue = cell.CellValue;
            var text = (cellValue == null) ? cell.InnerText : cellValue.Text;
            if ((cell.DataType != null) && (cell.DataType == CellValues.SharedString))
            {
                text = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>()
                                .ElementAt(Convert.ToInt32(cell.CellValue.Text))
                                .InnerText;
            }

            return (text ?? string.Empty).Trim();
        }
    }
}
