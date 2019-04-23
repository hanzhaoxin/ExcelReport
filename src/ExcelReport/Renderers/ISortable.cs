using ExcelReport.Contexts;

namespace ExcelReport.Renderers
{
    public interface ISortable
    {
        int SortNum(SheetContext sheetContext);
    }
}