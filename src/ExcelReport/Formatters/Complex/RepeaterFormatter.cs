/*
 类：RepeaterFormatter<TSource>、RepeaterFormatter<TThisSource, TSource>
 描述：重复单元格式化器、内嵌重复单元格式化器
 编 码 人：韩兆新 日期：2015年04月11日
 修改记录：
 * 2015年08月02日（韩兆新） 添加内嵌重复单元格式化器

*/
using System;
using System.Collections.Generic;
using System.Linq;
using ExcelReport.Exceptions;

namespace ExcelReport
{
    public class RepeaterFormatter<TSource> : ComplexFormatter<TSource>
    {
        #region 属性
        protected Parameter StartTagParameter { set; get; }
        protected Parameter EndTagParameter { set; get; }
        #endregion

        #region 0 构造函数
        /// <summary>
        /// 实例化重复单元格式化器
        /// </summary>
        /// <param name="startTagParameter">开始标签参数</param>
        /// <param name="endTagParameter">结束标签参数</param>
        /// <param name="dataSource">数据源</param>
        /// <param name="formatters">格式化器集合</param>
        public RepeaterFormatter(Parameter startTagParameter, Parameter endTagParameter,
            IEnumerable<TSource> dataSource, params EmbeddedFormatter<TSource>[] formatters)
            : base(dataSource, formatters)
        {
            StartTagParameter = startTagParameter;
            EndTagParameter = endTagParameter;
        }
        #endregion

        #region 1.0 格式化
        public override void Format(SheetAdapter sheetAdapter)
        {
            if (FormatterList.IsNullOrEmpty())
            {
                throw new ExcelReportFormatException("RepeaterFormatter is empty");
            }
            if (DataSource.IsNullOrEmpty())
            {
                sheetAdapter.RemoveRows(StartTagParameter.TagRowIndex, EndTagParameter.TagRowIndex); //删除模板行
            }
            else
            {
                for (int i = 0; i < DataSource.Count; i++)
                {
                    if (i < DataSource.Count - 1) //非最后一行数据时，复制模板
                    {
                        sheetAdapter.CopyRows(StartTagParameter.TagRowIndex, EndTagParameter.TagRowIndex, () =>
                        {
                            sheetAdapter.RemoveRow(StartTagParameter.TagRowIndex);
                            foreach (var formatter in FormatterList) //格式化行
                            {
                                formatter.Format(sheetAdapter, DataSource[i]);
                            }
                            sheetAdapter.RemoveRow(EndTagParameter.TagRowIndex);
                        });
                    }
                    else
                    {
                        sheetAdapter.RemoveRow(StartTagParameter.TagRowIndex);
                        foreach (var formatter in FormatterList) //格式化行
                        {
                            formatter.Format(sheetAdapter, DataSource[i]);
                        }
                        sheetAdapter.RemoveRow(EndTagParameter.TagRowIndex);
                    }

                }
            }
        }
        #endregion

        #region 1.1 添加内嵌格式化器
        /// <summary>
        /// 添加内嵌格式化器
        /// </summary>
        /// <param name="formatter">内嵌格式化器</param>
        /// <returns>当前对象</returns>
        public RepeaterFormatter<TSource> AddEmbeddedFormatter(EmbeddedFormatter<TSource> formatter)
        {
            this.FormatterList.Add(formatter);
            return this;
        }
        #endregion
    }


    public class RepeaterFormatter<TThisSource, TSource> : ComplexFormatter<TThisSource, TSource>
    {
        #region 属性
        protected Parameter StartTagParameter { set; get; }
        protected Parameter EndTagParameter { set; get; }
        #endregion

        #region 0 构造函数
        /// <summary>
        /// 实例化内嵌重复单元格式化器
        /// </summary>
        /// <param name="startTagParameter">开始标签参数</param>
        /// <param name="endTagParameter">结束标签参数</param>
        /// <param name="dgSetThisDataSource">赋值委托</param>
        /// <param name="formatters">格式化器集合</param>
        public RepeaterFormatter(Parameter startTagParameter, Parameter endTagParameter,
            Func<TSource, IEnumerable<TThisSource>> dgSetThisDataSource,
            params EmbeddedFormatter<TThisSource>[] formatters)
            : base(dgSetThisDataSource, formatters)
        {
            StartTagParameter = startTagParameter;
            EndTagParameter = endTagParameter;
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
            if (FormatterList.IsNullOrEmpty())
            {
                throw new ExcelReportFormatException("RepeaterFormatter is empty");
            }

            var thisDataSource = DgSetThisDataSource(dataSource).ToList();
            if (thisDataSource.IsNullOrEmpty())
            {
                sheetAdapter.RemoveRows(StartTagParameter.TagRowIndex, EndTagParameter.TagRowIndex); //删除模板行
            }
            else
            {
                for (int i = 0; i < thisDataSource.Count; i++)
                {
                    if (i < thisDataSource.Count - 1) //非最后一行数据时，复制模板
                    {
                        sheetAdapter.CopyRows(StartTagParameter.TagRowIndex, EndTagParameter.TagRowIndex, () =>
                        {
                            sheetAdapter.RemoveRow(StartTagParameter.TagRowIndex);
                            foreach (var formatter in FormatterList) //格式化行
                            {
                                formatter.Format(sheetAdapter, thisDataSource[i]);
                            }
                            sheetAdapter.RemoveRow(EndTagParameter.TagRowIndex);
                        });
                    }
                    else
                    {
                        sheetAdapter.RemoveRow(StartTagParameter.TagRowIndex);
                        foreach (var formatter in FormatterList) //格式化行
                        {
                            formatter.Format(sheetAdapter, thisDataSource[i]);
                        }
                        sheetAdapter.RemoveRow(EndTagParameter.TagRowIndex);
                    }
                }
            }
        }
        #endregion

        #region 1.1 添加内嵌格式化器
        /// <summary>
        /// 添加内嵌格式化器
        /// </summary>
        /// <param name="formatter">内嵌格式化器</param>
        /// <returns>当前格式化器</returns>
        public RepeaterFormatter<TThisSource, TSource> AddEmbeddedFormatter(EmbeddedFormatter<TThisSource> formatter)
        {
            this.FormatterList.Add(formatter);
            return this;
        }
        #endregion
    }
}