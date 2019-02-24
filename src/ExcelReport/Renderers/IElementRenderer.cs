namespace ExcelReport.Renderers
{
    public interface IElementRenderer
    {
        void Render(SheetAdapter sheetAdapter);
    }
}