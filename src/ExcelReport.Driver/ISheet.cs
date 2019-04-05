using System.Collections.Generic;

namespace ExcelReport.Driver
{
    public interface ISheet : IEnumerable<IRow>, IAdapter
    {
        string SheetName { get; }

        IRow this[int rowIndex]
        {
            get;
        }

        int CopyRows(int start, int end);

        int RemoveRows(int start, int end);
    }
}