/*
 类：TableColumnInfo
 描述：Table元素的列信息
 编 码 人：韩兆新 日期：2015年01月17日
 修改记录：

*/

using System;

namespace ExcelReport
{
    public class TableColumnInfo<TSource>
    {
        #region 成员字段及属性

        private int _columnIndex;
        private Func<TSource, object> _dgSetValue;

        public int ColumnIndex
        {
            get { return _columnIndex; }
        }

        public Func<TSource, object> DgSetValue
        {
            get { return _dgSetValue; }
        }

        #endregion 成员字段及属性

        ///  构造函数
        /// <param name="columnIndex"></param>
        /// <param name="dgSetValue"></param>
        public TableColumnInfo(int columnIndex, Func<TSource, object> dgSetValue)
        {
            _columnIndex = columnIndex;
            _dgSetValue = dgSetValue;
        }

    }
}