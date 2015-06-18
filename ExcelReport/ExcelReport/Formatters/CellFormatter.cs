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

        protected Point CellPoint
        {
            get { return _cellPoint; }
        }

        protected object Value
        {
            get { return _value; }
        }

        #endregion 成员字段及属性

        /// 构造函数
        /// <param name="cellPoint"></param>
        /// <param name="value"></param>
        public CellFormatter(Point cellPoint, object value)
        {
            _cellPoint = cellPoint;
            _value = value;
        }

        /// 构造函数
        /// <param name="rowIndex"></param>
        /// <param name="columnIndex"></param>
        /// <param name="value"></param>
        public CellFormatter(int rowIndex, int columnIndex, object value)
        {
            _cellPoint = new Point(rowIndex, columnIndex);
            _value = value;
        }


        /// 格式化操作
        /// <param name="context"></param>
        public override void Format(SheetFormatterContext context)
        {
            var rowIndex = context.GetCurrentRowIndex(CellPoint.X);
            var row = context.Sheet.GetRow(rowIndex);
            if (null == row)
            {
                row = context.Sheet.CreateRow(rowIndex);
            }
            var cell = row.GetCell(CellPoint.Y);
            if (null == cell)
            {
                cell = row.CreateCell(CellPoint.Y);
            }
            SetCellValue(cell, Value);
        }

    }
}