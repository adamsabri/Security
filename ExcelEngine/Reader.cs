using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelEngine
{
    public class Reader : IReader
    {
        private readonly IReader _reader;

        public string FilePath { get; private set; }

        public Reader(string filePath, XlsxReaderType type = XlsxReaderType.Epplus)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath)) return;
            FilePath = filePath;
            if (filePath.EndsWith(".xls"))
                _reader = new NPOI.XlsReader(filePath);
            else if (filePath.EndsWith(".xlsx"))
            {
                switch (type)
                {
                    case XlsxReaderType.Npoi:
                        _reader = new NPOI.XlsxReader(filePath);
                        break;
                    case XlsxReaderType.Ooxml:
                        _reader = new OOXML.XlsxReader(filePath, OooxmlAproach.Sax);
                        break;
                    default:
                        _reader = new EPPlus.XlsxReader(filePath);
                        break;
                }
            }
        }

        public IList<string> ConvertToList()
        {
           return _reader.ConvertToList();
        }
    }

    public enum XlsxReaderType
    {
        Epplus = 0,
        Npoi = 1,
        Ooxml = 2,
    }

    public enum OooxmlAproach
    {
        /// <summary>
        /// The DOM approach.
        /// The method will definitely work for reading cell values within a worksheet.
        /// However, if the worksheet is quite large then the program's memory footprint will
        /// also be quite large. In fact, you are left at the mercy of the garbage collector,
        /// which may result in the program throwing an Out of Memory exception. 
        /// </summary>
        Dom = 0,
        /// <summary>
        /// The SAX approach.
        /// The method takes advantage of the OpenXmlReader class,
        /// which allows you to read an entire part without loading the part in memory.
        /// </summary>
        Sax = 1,
    }
}
