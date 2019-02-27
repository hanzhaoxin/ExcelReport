using ExcelReport.Contexts;

namespace ExcelReport.Renderers
{
    public interface IElementRenderer
    {
        void Render(SheetContext sheetContext);
    }
}