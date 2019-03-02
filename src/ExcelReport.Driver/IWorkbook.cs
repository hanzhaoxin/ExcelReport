using System.Collections.Generic;

namespace ExcelReport.Driver
{
    public interface IWorkbook : IEnumerable<ISheet>
    {
        ISheet this[string sheetName]
        {
            get;
        }

        byte[] SaveToBuffer();
    }
}