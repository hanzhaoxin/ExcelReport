/*
 类：CellFormatter
 描述：单元格（元素）格式化器
 编 码 人：韩兆新 日期：2015年01月17日
 修改记录：

*/

using System.Drawing;

namespace ExcelReport
{
    public class CellFormatter : ElementFormatter
    {
        #region 成员字段及属性

        private Point _cellPoint;
        private object _value;

        #endregion 成员字段及属性

        #region 构造函数

        public CellFormatter(Point cellPoint, object value)
        {
            _cellPoint = cellPoint;
            _value = value;
        }

        public CellFormatter(int rowIndex, int columnIndex, object value)
        {
            _cellPoint = new Point(rowIndex, columnIndex);
            _value = value;
        }

        #endregion 构造函数

        #region 格式化操作

        public override void Format(SheetFormatterContext context)
        {
            var rowIndex = context.GetCurrentRowIndex(_cellPoint.X);
            var row = context.Sheet.GetRow(rowIndex);
            if (null == row)
            {
                row = context.Sheet.CreateRow(rowIndex);
            }
            var cell = row.GetCell(_cellPoint.Y);
            if (null == cell)
            {
                cell = row.CreateCell(_cellPoint.Y);
            }
            SetCellValue(cell, _value);
        }

        #endregion 格式化操作
    }
}