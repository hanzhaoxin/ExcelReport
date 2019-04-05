using System.Collections.Generic;

namespace ExcelReport.Driver
{
    public interface IWorkbook : IEnumerable<ISheet>, IAdapter
    {
        ISheet this[string sheetName]
        {
            get;
        }

        byte[] SaveToBuffer();
    }
}