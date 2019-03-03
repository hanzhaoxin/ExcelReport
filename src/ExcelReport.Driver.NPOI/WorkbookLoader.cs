using ExcelReport.Driver.NPOI.Extends;
using NPOI.Extend;

namespace ExcelReport.Driver.NPOI
{
    public class WorkbookLoader : IWorkbookLoader
    {
        public IWorkbook Load(string filePath)
        {
            return NPOIHelper.LoadWorkbook(filePath).GetAdapter();
        }
    }
}