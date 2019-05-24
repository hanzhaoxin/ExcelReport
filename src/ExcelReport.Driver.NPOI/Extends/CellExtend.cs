using NpoiCell = NPOI.SS.UserModel.ICell;

namespace ExcelReport.Driver.NPOI.Extends
{
    internal static class CellExtend
    {
        public static Cell GetAdapter(this NpoiCell cell)
        {
            if (null == cell)
            {
                return null;
            }
            return new Cell(cell);
        }
    }
}