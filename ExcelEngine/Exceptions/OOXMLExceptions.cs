using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelEngine.Exceptions
{
    public class OoxmlInitializeException : Exception
    {
        public OoxmlInitializeException(string message)
            : base(message)
        {
        }
        public OoxmlInitializeException(string message, Exception e)
            : base(message, e)
        {
        }
    }

    public class OoxmlXlsxDomReaderException : Exception
    {
        public OoxmlXlsxDomReaderException(string message)
            : base(message)
        {
        }
        public OoxmlXlsxDomReaderException(string message, Exception e)
            : base(message, e)
        {
        }
    }

    public class OoxmlXlsxSaxReaderException : Exception
    {
        public OoxmlXlsxSaxReaderException(string message)
            : base(message)
        {
        }
        public OoxmlXlsxSaxReaderException(string message, Exception e)
            : base(message, e)
        {
        }
    }
}
