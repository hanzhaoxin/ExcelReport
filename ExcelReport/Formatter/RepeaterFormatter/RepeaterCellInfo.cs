/*
 类：RepeaterCellInfo
 描述：Repeater元素的单元格信息
 编 码 人：韩兆新 日期：2015年04月11日
 修改记录：

*/
using System;
using System.Drawing;

namespace ExcelReport
{
    public class RepeaterCellInfo<TSource>
    {
        #region 成员字段及属性

        private Point _cellPoint;
        private Func<TSource, object> _dgSetValue;

        public Point CellPoint
        {
            get { return _cellPoint; }
        }
        public Func<TSource, object> DgSetValue
        {
            get { return _dgSetValue; }
        } 

        #endregion 成员字段及属性

        #region 构造函数

        public RepeaterCellInfo(Point cellPoint, Func<TSource, object> dgSetValue)
        {
            _cellPoint = cellPoint;
            _dgSetValue = dgSetValue;
        }

        #endregion 构造函数
    }
}
