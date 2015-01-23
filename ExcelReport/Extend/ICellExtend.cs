/*
 类：ICellExtend
 描述：为ICell扩展方法
 编 码 人：韩兆新 日期：2015年01月23日
 修改记录：

*/
using NPOI.SS.UserModel;

namespace ExcelReport
{
    static class ICellExtend
    {
        public static CellSpan GetSpan(this ICell cell)
        {
            var cellSpan = new CellSpan(1,1);
            if (cell.IsMergedCell)
            {
                int regionsNum = cell.Sheet.NumMergedRegions;
                for (int i = 0; i < regionsNum; i++)
                {
                    var range = cell.Sheet.GetMergedRegion(i);
                    if (range.FirstRow == cell.RowIndex && range.FirstColumn == cell.ColumnIndex)
                    {
                        cellSpan.RowSpan = range.LastRow - range.FirstRow + 1;
                        cellSpan.ColSpan = range.LastColumn - range.FirstColumn + 1;
                        break;
                    }
                }
            }
            return cellSpan;
        }
    }
}
