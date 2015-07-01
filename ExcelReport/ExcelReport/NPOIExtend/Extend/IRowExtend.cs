/*
 类：IRowExtend
 描述：IRow扩展方法
 编 码 人：韩兆新 日期：2015年05月15日
 修改记录：

*/
using System;
using NPOI.SS.UserModel;

namespace ExcelReport
{
    internal static class IRowExtend
    {
        /// <summary>
        /// 复制源行到目标行(忽略合并单元格等信息)
        /// </summary>
        /// <param name="sourceRow"></param>
        /// <param name="targetRowIndex"></param>
        /// <returns></returns>
        public static IRow CopyTo(this IRow sourceRow, int targetRowIndex)
        {
            if (targetRowIndex == sourceRow.RowNum)
            {
                throw new ArgumentException("目标行标不能等于源行标。");
            }
            ISheet currentSheet = sourceRow.Sheet;
            IRow targetRow = currentSheet.GetRow(targetRowIndex);
            if (targetRow != null)
            {
                currentSheet.ShiftRows(targetRowIndex, currentSheet.LastRowNum, 1);
            }
            else
            {
                targetRow = currentSheet.CreateRow(targetRowIndex);
            }
            targetRow.Height = sourceRow.Height;
            targetRow.ZeroHeight = sourceRow.ZeroHeight;

            #region 复制单元格
            foreach (var sourceCell in sourceRow.Cells)
            {
                ICell targetCell = targetRow.GetCell(sourceCell.ColumnIndex);
                if (null == targetCell)
                {
                    targetCell = targetRow.CreateCell(sourceCell.ColumnIndex);
                }
                if (null != sourceCell.CellStyle) targetCell.CellStyle = sourceCell.CellStyle;
                if (null != sourceCell.CellComment) targetCell.CellComment = sourceCell.CellComment;
                if (null != sourceCell.Hyperlink) targetCell.Hyperlink = sourceCell.Hyperlink;
                var cfrs = sourceCell.GetConditionalFormattingRules();  //复制条件样式
                if (null != cfrs && cfrs.Length > 0)
                {
                    targetCell.AddConditionalFormattingRules(cfrs);
                }
                targetCell.SetCellType(sourceCell.CellType);
                #region 复制值
                switch (sourceCell.CellType)
                {
                    case CellType.Numeric:
                        targetCell.SetCellValue(sourceCell.NumericCellValue);
                        break;
                    case CellType.String:
                        targetCell.SetCellValue(sourceCell.RichStringCellValue);
                        break;
                    case CellType.Formula:
                        targetCell.SetCellFormula(sourceCell.CellFormula);
                        break;
                    case CellType.Blank:
                        targetCell.SetCellValue(sourceCell.StringCellValue);
                        break;
                    case CellType.Boolean:
                        targetCell.SetCellValue(sourceCell.BooleanCellValue);
                        break;
                    case CellType.Error:
                        targetCell.SetCellErrorValue(sourceCell.ErrorCellValue);
                        break;
                }
                #endregion
            }
            #endregion

            return targetRow;
        }
    }
}
