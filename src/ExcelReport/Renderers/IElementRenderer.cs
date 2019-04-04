using ExcelReport.Contexts;

namespace ExcelReport.Renderers
{
    public interface IElementRenderer : ISortable
    {
        void Render(SheetContext sheetContext);
    }
}