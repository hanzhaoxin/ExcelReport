namespace ExcelReport.Driver
{
    public interface IWorkbookLoader
    {
        IWorkbook Load(string filePath);
    }
}