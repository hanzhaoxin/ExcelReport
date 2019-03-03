using NpoiCell = NPOI.SS.UserModel.ICell;

namespace ExcelReport.Driver.NPOI.Extends
{
    internal static class CellExtend
    {
        public static Cell GetAdapter(this NpoiCell cell)
        {
            return new Cell(cell);
        }
    }
}