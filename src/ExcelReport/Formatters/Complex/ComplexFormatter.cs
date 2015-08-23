/*
 类：ComplexFormatter<TSource>、ComplexFormatter<TThisSource, TSource>
 描述：复合格式化器、内嵌复合格式化器
 编 码 人：韩兆新  日期：2015年08月02日
 修改记录：
 *
 
*/
using System;
using System.Collections.Generic;

namespace ExcelReport
{
    public abstract class ComplexFormatter<TSource> : ElementFormatter
    {
        #region 属性
        protected IList<EmbeddedFormatter<TSource>> FormatterList { set; get; }
        protected IList<TSource> DataSource { set; get; }
        #endregion

        #region 0 构造函数
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="dataSource">数据源</param>
        /// <param name="formatters">格式化器集合</param>
        protected ComplexFormatter(IEnumerable<TSource> dataSource, IEnumerable<EmbeddedFormatter<TSource>> formatters)
        {
            DataSource = new List<TSource>(dataSource);
            FormatterList = new List<EmbeddedFormatter<TSource>>(formatters);
        }
        #endregion

       
    }

    public abstract class ComplexFormatter<TThisSource, TSource> : EmbeddedFormatter<TSource>
    {
        #region 属性
        protected IList<EmbeddedFormatter<TThisSource>> FormatterList { set; get; }
        protected Func<TSource, IEnumerable<TThisSource>> DgSetThisDataSource { set; get; }
        #endregion

        #region 0 构造函数
        /// <summary>
        /// 构成函数
        /// </summary>
        /// <param name="dgSetThisDataSource">赋值委托</param>
        /// <param name="formatters">格式化器集合</param>
        protected ComplexFormatter(Func<TSource, IEnumerable<TThisSource>> dgSetThisDataSource,
            IEnumerable<EmbeddedFormatter<TThisSource>> formatters)
        {
            DgSetThisDataSource = dgSetThisDataSource;
            FormatterList = new List<EmbeddedFormatter<TThisSource>>(formatters);
        }
        #endregion
    }
}