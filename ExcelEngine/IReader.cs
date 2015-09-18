using System.Collections.Generic;

namespace ExcelEngine
{
    public interface IReader
    {
        string FilePath { get; }
        IList<string> ConvertToList();
    }
}
