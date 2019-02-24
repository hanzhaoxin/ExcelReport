namespace ExcelReport.Renderers
{
    public interface IEmbeddedRenderer<TSource>
    {
        void Render(SheetAdapter sheetAdapter, TSource dataSource);
    }
}