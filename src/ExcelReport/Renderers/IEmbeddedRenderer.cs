using ExcelReport.Contexts;

namespace ExcelReport.Renderers
{
    public interface IEmbeddedRenderer<TSource> : ISortable
    {
        void Render(SheetContext sheetContext, TSource dataSource);
    }
}