/*
 类：SheetFormatterContext
 描述：Sheet格式化的上下文
 编 码 人：韩兆新 日期：2015年01月17日
 修改记录：

*/

using System.Collections.Generic;
using NPOI.SS.UserModel;

namespace ExcelReport
{
    public class SheetFormatterContext
    {
        #region 成员字段及属性

        private int _increaseRowsCount = 0;

        public ISheet Sheet { get; set; }

        public IEnumerable<ElementFormatter> Formatters { get; set; }

        #endregion 成员字段及属性

        #region 构造函数

        public SheetFormatterContext()
        {
        }

        public SheetFormatterContext(ISheet sheet, IEnumerable<ElementFormatter> formatters)
        {
            this.Sheet = sheet;
            this.Formatters = formatters;
        }

        #endregion 构造函数

        #region 获取指定行当前行标

        /// <summary>
        /// 获取指定行当前行标
        /// </summary>
        /// <param name="rowIndexInTemplate">指定行在模板中的行标</param>
        /// <returns>当前行标</returns>
        public int GetCurrentRowIndex(int rowIndexInTemplate)
        {
            return rowIndexInTemplate + _increaseRowsCount;
        }

        #endregion 获取指定行当前行标

        #region 在指定行后插入一行（并将指定行作为模板复制样式）

        /// <summary>
        /// 在指定行后插入一行（并将指定行作为模板复制样式）
        /// </summary>
        /// <param name="templateRowIndex">模板行在模板中的行标</param>
        public void InsertEmptyRow(int templateRowIndex)
        {
            var templateRow = Sheet.GetRow(GetCurrentRowIndex(templateRowIndex));
            var insertRowIndex = GetCurrentRowIndex(templateRowIndex + 1);
            if (insertRowIndex < Sheet.LastRowNum)
            {
                Sheet.ShiftRows(insertRowIndex, Sheet.LastRowNum, 1, true, false);
            }
            var newRow = Sheet.CreateRow(GetCurrentRowIndex(templateRowIndex + 1));
            newRow.Height = templateRow.Height;
            _increaseRowsCount++;
            foreach (var cell in templateRow.Cells)
            {
                newRow.CreateCell(cell.ColumnIndex).CellStyle = cell.CellStyle;
            }
        }

        #endregion 在指定行后插入一行（并将指定行作为模板复制样式）

        #region 清除指定行单元格内容

        /// <summary>
        /// 清除指定行单元格内容
        /// </summary>
        /// <param name="rowIndex">指定行在模板中的行标</param>
        public void ClearRowContent(int rowIndex)
        {
            var row = Sheet.GetRow(GetCurrentRowIndex(rowIndex));
            foreach (var cell in row.Cells)
            {
                cell.SetCellValue(string.Empty);
            }
        }

        #endregion 清除指定行单元格内容

        #region 删除指定行

        /// <summary>
        /// 删除指定行
        /// </summary>
        /// <param name="rowIndex">指定行在模板中的行标</param>
        public void RemoveRow(int rowIndex)
        {
            var row = Sheet.GetRow(GetCurrentRowIndex(rowIndex));
            Sheet.RemoveRow(row);
        }

        #endregion 删除指定行

        #region 格式化Sheet

        /// <summary>
        /// 格式化Sheet
        /// </summary>
        public void Format()
        {
            if (null == Sheet || null == Formatters)
            {
                return;
            }
            foreach (var formatter in Formatters)
            {
                formatter.Format(this);
            }
        }

        #endregion 格式化Sheet
    }
}