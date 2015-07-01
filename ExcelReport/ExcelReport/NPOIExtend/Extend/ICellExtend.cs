/*
 类：ICellExtend
 描述：ICell扩展方法
 编 码 人：韩兆新 日期：2015年01月17日
 修改记录：

*/

using System.Collections.Generic;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace ExcelReport
{
    internal static class ICellExtend
    {
        /// 获取cell的CellSpan信息
        /// <param name="cell"></param>
        /// <returns></returns>
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

        #region 条件格式
        
        /// 获取条件格式规则
        /// <param name="cell">单元格</param>
        /// <returns></returns>
        public static IConditionalFormattingRule[] GetConditionalFormattingRules(this ICell cell)
        {
            var cfrList = new List<IConditionalFormattingRule>();

            var scf = cell.Sheet.SheetConditionalFormatting;
            for (int i = 0; i < scf.NumConditionalFormattings; i++)
            {
                IConditionalFormatting cf = scf.GetConditionalFormattingAt(i);
                if (cell.ExistConditionalFormatting(cf))
                {
                    for (int j = 0; j < cf.NumberOfRules; j++)
                    {
                        cfrList.Add(cf.GetRule(j));
                    }
                }
            }
            return cfrList.ToArray();
        }

       
        /// 添加条件格式规则
        /// <param name="cell">单元格</param>
        /// <param name="cfrs">条件格式规则</param>
        public static void AddConditionalFormattingRules(this ICell cell, IConditionalFormattingRule[] cfrs)
        {
            CellRangeAddress[] regions =
            {
                new CellRangeAddress(cell.RowIndex, cell.RowIndex, cell.ColumnIndex, cell.ColumnIndex)
            };
            cell.Sheet.SheetConditionalFormatting.AddConditionalFormatting(regions, cfrs);
        }

        /// 单元格是否存在条件格式
        /// <param name="cell"></param>
        /// <param name="cf"></param>
        /// <returns></returns>
        private static bool ExistConditionalFormatting(this ICell cell, IConditionalFormatting cf)
        {
            var cfRangeAddrs = cf.GetFormattingRanges();
            foreach (var cfRangeAddr in cfRangeAddrs)
            {
                if (cell.RowIndex >= cfRangeAddr.FirstRow && cell.RowIndex <= cfRangeAddr.LastRow
                    && cell.ColumnIndex >= cfRangeAddr.FirstColumn && cell.ColumnIndex <= cfRangeAddr.LastColumn)
                {
                    return true;
                }
            }
            return false;
        }

        #endregion
    }
}
