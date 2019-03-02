using ExcelReport.Driver.NPOIDriver.Extends;
using NPOI.Extend;

namespace ExcelReport.Driver.NPOIDriver
{
    public class WorkbookLoader : IWorkbookLoader
    {
        public IWorkbook Load(string filePath)
        {
            return NPOIHelper.LoadWorkbook(filePath).GetAdapter();
        }
    }
}