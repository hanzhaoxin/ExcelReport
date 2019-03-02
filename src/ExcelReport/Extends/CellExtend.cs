using ExcelReport.Driver;

namespace ExcelReport.Extends
{
    public static class CellExtend
    {
        public static string GetStringValue(this ICell cell)
        {
            if (cell.IsNull())
            {
                return string.Empty;
            }

            try
            {
                return cell.Value.CastTo<string>();
            }
            catch
            {
                return string.Empty;
            }
        }
    }
}