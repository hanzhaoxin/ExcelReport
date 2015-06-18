/*
 类：PartFormatter
 描述：单元格局部（元素）格式化器
 编 码 人：韩兆新 日期：2015年01月25日
 修改记录：

*/

using System.Drawing;
using NPOI.SS.UserModel;

namespace ExcelReport
{
    public class PartFormatter:ElementFormatter
    {
        #region 成员字段
        private Point _cellPoint;
        private string _parameterName;
        private string _value;
        #endregion

        #region 属性
        protected Point CellPoint
        {
            get { return _cellPoint; }
        }

        protected string ParameterName
        {
            get { return _parameterName; }
        }

        protected string Value
        {
            get { return _value; }
        }
        #endregion

        /// 构造函数
        /// <param name="cellPoint"></param>
        /// <param name="parameterName"></param>
        /// <param name="value"></param>
        public PartFormatter(Point cellPoint, string parameterName ,string value)
        {
            this._cellPoint = cellPoint;
            this._parameterName = parameterName;
            this._value = value;
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
            if (cell.CellType.Equals(CellType.String))
            {
                SetCellValue(cell, cell.StringCellValue.Replace(string.Format("$[{0}]", ParameterName), Value));
            }
        }
    }
}
