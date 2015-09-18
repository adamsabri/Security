using System;

namespace ExcelEngine.Exceptions
{
    public class NpoiInitializeException : Exception
    {
        public NpoiInitializeException(string message) : base(message)
        {
        }
        public NpoiInitializeException(string message ,Exception e)
            : base(message, e)
        {
        }
    }

    public class NpoiXlsxReaderException : Exception
    {
        public NpoiXlsxReaderException(string message)
            : base(message)
        {
        }
        public NpoiXlsxReaderException(string message, Exception e)
            : base(message, e)
        {
        }
    }

    public class NpoiXlsReaderException : Exception
    {
        public NpoiXlsReaderException(string message)
            : base(message)
        {
        }
        public NpoiXlsReaderException(string message, Exception e)
            : base(message, e)
        {
        }
    }
}
