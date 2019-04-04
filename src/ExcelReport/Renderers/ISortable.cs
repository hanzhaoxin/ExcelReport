using ExcelReport.Contexts;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelReport.Renderers
{
    public interface ISortable
    {
        int SortNum(SheetContext sheetContext);
    }
}
