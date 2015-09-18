using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;

namespace ExcelEngine.NPOI
{
    public class Writer : BaseWriter
    {
        private static IWriter _writer;
        private static Writer _instance;

        public static Writer Instance
        {
            get { return _instance ?? (_instance = new Writer()); }
        }

        #region IWriter Implementation

        public override void CreateWorkBook(string destinationPath)
        {
            if (destinationPath.EndsWith(".xls"))
                _writer = new XlsWriter();
            else if (destinationPath.EndsWith(".xlsx"))
                _writer = new XlsxWriter();
            else
                return;

            _writer.CreateWorkBook(destinationPath);
        }

        #endregion

    }
}
