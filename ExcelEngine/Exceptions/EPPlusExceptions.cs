using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelEngine.Exceptions
{
    public class EpplusInitializeException : Exception
    {
        public EpplusInitializeException(string message)
            : base(message)
        {
        }
        public EpplusInitializeException(string message, Exception e)
            : base(message, e)
        {
        }
    }

    public class EpplusXlsxReaderException : Exception
    {
        public EpplusXlsxReaderException(string message)
            : base(message)
        {
        }
        public EpplusXlsxReaderException(string message, Exception e)
            : base(message, e)
        {
        }
    }
}
