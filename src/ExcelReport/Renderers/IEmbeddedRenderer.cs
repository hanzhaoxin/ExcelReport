using ExcelReport.Contexts;

namespace ExcelReport.Renderers
{
    public interface IEmbeddedRenderer<TSource>
    {
        void Render(SheetContext sheetContext, TSource dataSource);
    }
}