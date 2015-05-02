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
        private Point _cellPoint;
        private string _parameterName;
        private string _value;

        public PartFormatter(Point cellPoint, string parameterName ,string value)
        {
            this._cellPoint = cellPoint;
            this._parameterName = parameterName;
            this._value = value;
        }

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
            if (cell.CellType.Equals(CellType.String))
            {
                SetCellValue(cell, cell.StringCellValue.Replace(string.Format("$[{0}]", _parameterName), _value));
            }
        }
    }
}
