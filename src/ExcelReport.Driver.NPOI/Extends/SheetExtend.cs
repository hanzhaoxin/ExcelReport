using NpoiSheet = NPOI.SS.UserModel.ISheet;

namespace ExcelReport.Driver.NPOI.Extends
{
    internal static class SheetExtend
    {
        public static Sheet GetAdapter(this NpoiSheet sheet)
        {
            return new Sheet(sheet);
        }
    }
}