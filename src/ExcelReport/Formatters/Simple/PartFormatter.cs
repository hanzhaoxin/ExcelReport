/*
 类：PartFormatter、PartFormatter<TSource>
 描述：局部格式化器、内嵌局部格式化器
 编 码 人：韩兆新 日期：2015年01月25日
 修改记录：
 * 2015年08月01日（韩兆新） 添加内嵌局部格式化器
 
*/
using System;
using ExcelReport.Exceptions;
using NPOI.SS.UserModel;
using NPOI.Extend;

namespace ExcelReport
{
    public class PartFormatter : SimpleFormatter
    {
        #region 0 构造函数
        /// <summary>
        ///     实例化部分格式化器
        /// </summary>
        /// <param name="parameter">参数</param>
        /// <param name="value">值</param>
        public PartFormatter(Parameter parameter, object value)
            : base(parameter, value)
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
                throw new ExcelReportException("row is null");
            }
            ICell cell = row.GetCell(Parameter.ColumnIndex);
            if (null == cell)
            {
                throw new ExcelReportException("cell is null");
            }

            cell.SetValue(cell.StringCellValue.Replace(string.Format("$[{0}]", Parameter.Name), Value.CastTo<string>()));
        }
        #endregion
    }

    public class PartFormatter<TSource> : SimpleFormatter<TSource>
    {
        #region 0 构造函数
        /// <summary>
        /// 实例化内嵌部分格式化器
        /// </summary>
        /// <param name="parameter">参数</param>
        /// <param name="dgSetValue">赋值委托</param>
        public PartFormatter(Parameter parameter, Func<TSource, object> dgSetValue)
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
                throw new ExcelReportException("row is null");
            }
            ICell cell = row.GetCell(Parameter.ColumnIndex);
            if (null == cell)
            {
                throw new ExcelReportException("cell is null");
            }

            cell.SetValue(cell.StringCellValue.Replace(string.Format("$[{0}]", Parameter.Name),
                DgSetValue(dataSource).CastTo<string>()));
        }
        #endregion
    }
}