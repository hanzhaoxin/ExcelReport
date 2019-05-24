using NpoiRow = NPOI.SS.UserModel.IRow;

namespace ExcelReport.Driver.NPOI.Extends
{
    internal static class RowExtend
    {
        public static Row GetAdapter(this NpoiRow row)
        {
            if (null == row)
            {
                return null;
            }
            return new Row(row);
        }
    }
}