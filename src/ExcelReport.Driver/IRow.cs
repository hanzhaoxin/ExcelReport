using System.Collections.Generic;

namespace ExcelReport.Driver
{
    public interface IRow : IEnumerable<ICell>
    {
        ICell this[int columnIndex]
        {
            get;
        }
    }
}