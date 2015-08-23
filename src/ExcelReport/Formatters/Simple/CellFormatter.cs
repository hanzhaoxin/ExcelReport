/*
 类：CellFormatter、CellFormatter<TSource>
 描述：单元格格式化器、内嵌单元格格式化器
 编 码 人：韩兆新 日期：2015年01月17日
 修改记录：
 * 2015年08月01日（韩兆新） 添加内嵌单元格格式化器

*/
using System;
using ExcelReport.Exceptions;
using NPOI.SS.UserModel;
using NPOI.Extend;

namespace ExcelReport
{
    public class CellFormatter : SimpleFormatter
    {
        #region 0 构造函数

        /// <summary>
        /// 实例化单元格格式化器
        /// </summary>
        /// <param name="parameter">参数</param>
        /// <param name="value">值</param>
        public CellFormatter(Parameter parameter, object value) : base(parameter, value)
        {
        }

        #endregion

        #region 1.0 格式化
        /// <summary>
        ///     格式化
        /// </summary>
        /// <param name="sheetAdapter">Sheet适配器</param>
        public override void Format(SheetAdapter sheetAdapter)
        {
            IRow row = sheetAdapter.GetRow(Parameter.RowIndex);
            if (null == row)
            {
                throw new ExcelReportFormatException("row is null");
            }
            ICell cell = row.GetCell(Parameter.ColumnIndex);
            if (null == cell)
            {
                throw new ExcelReportFormatException("cell is null");
            }
            cell.SetValue(Value);
        }
        #endregion
    }

    public class CellFormatter<TSource> : SimpleFormatter<TSource>
    {
        #region 0 构造函数
        /// <summary>
        /// 实例化内嵌单元格格式化器
        /// </summary>
        /// <param name="parameter">参数</param>
        /// <param name="dgSetValue">赋值委托</param>
        public CellFormatter(Parameter parameter, Func<TSource, object> dgSetValue)
            : base(parameter, dgSetValue)
        {
        }
        #endregion

        #region 1.0 格式化
        /// <summary>
        ///     格式化
        /// </summary>
        /// <param name="sheetAdapter">Sheet适配器</param>
        /// <param name="dataSource">数据源</param>
        public override void Format(SheetAdapter sheetAdapter, TSource dataSource)
        {
            IRow row = sheetAdapter.GetRow(Parameter.RowIndex);
            if (null == row)
            {
                throw new ExcelReportFormatException("row is null");
            }
            ICell cell = row.GetCell(Parameter.ColumnIndex);
            if (null == cell)
            {
                throw new ExcelReportFormatException("cell is null");
            }
            cell.SetValue(DgSetValue(dataSource));
        }
        #endregion
    }
}