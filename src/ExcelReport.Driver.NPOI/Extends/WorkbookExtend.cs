using NpoiWorkbook = NPOI.SS.UserModel.IWorkbook;

namespace ExcelReport.Driver.NPOI.Extends
{
    internal static class WorkbookExtend
    {
        public static Workbook GetAdapter(this NpoiWorkbook workbook)
        {
            if (null == workbook)
            {
                return null;
            }
            return new Workbook(workbook);
        }
    }
}