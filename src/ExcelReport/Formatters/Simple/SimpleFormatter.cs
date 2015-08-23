/*
 类：SimpleFormatter、SimpleFormatter<TSource>
 描述：简单格式化器、内嵌简单格式化器
 编 码 人：韩兆新  日期：2015年08月01日
 修改记录：
 *
 
*/
using System;

namespace ExcelReport
{
    public abstract class SimpleFormatter : ElementFormatter
    {
        #region 属性
        protected Parameter Parameter { set; get; }
        protected object Value { set; get; }
        #endregion

        #region 0 构造函数
        /// <summary>
        ///     构造函数
        /// </summary>
        /// <param name="parameter">参数</param>
        /// <param name="value">值</param>
        protected SimpleFormatter(Parameter parameter, object value)
        {
            Parameter = parameter;
            Value = value;
        }
        #endregion
    }

    public abstract class SimpleFormatter<TSource> : EmbeddedFormatter<TSource>
    {
        #region 属性
        protected Parameter Parameter { set; get; }
        protected Func<TSource, object> DgSetValue { set; get; }
        #endregion

        #region 0 构造函数
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="parameter">参数</param>
        /// <param name="dgSetValue">赋值委托</param>
        protected SimpleFormatter(Parameter parameter, Func<TSource, object> dgSetValue)
        {
            Parameter = parameter;
            DgSetValue = dgSetValue;
        }
        #endregion
    }
}