using NpoiWorkbook = NPOI.SS.UserModel.IWorkbook;

namespace ExcelReport.Driver.NPOIDriver.Extends
{
    internal static class WorkbookExtend
    {
        public static Workbook GetAdapter(this NpoiWorkbook workbook)
        {
            return new Workbook(workbook);
        }
    }
}