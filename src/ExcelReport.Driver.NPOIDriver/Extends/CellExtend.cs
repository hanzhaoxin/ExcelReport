using NpoiCell = NPOI.SS.UserModel.ICell;

namespace ExcelReport.Driver.NPOIDriver.Extends
{
    internal static class CellExtend
    {
        public static Cell GetAdapter(this NpoiCell cell)
        {
            return new Cell(cell);
        }
    }
}