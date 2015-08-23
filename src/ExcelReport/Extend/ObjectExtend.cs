/*
 类：ObjectExtend
 描述：object的扩展方法
 编 码 人：韩兆新 日期：2015年08月01日
 修改记录：

*/
using System;

namespace ExcelReport
{
    public static class ObjectExtend
    {
        /// <summary>
        ///     判断是否为空
        /// </summary>
        /// <param name="value">要判断的值</param>
        /// <returns>是返回True，不是返回False</returns>
        public static bool IsNull(this object value)
        {
            return null == value;
        }

        /// <summary>
        ///     把对象类型转换为指定类型
        /// </summary>
        /// <param name="value">要转换的值</param>
        /// <param name="targetType">目标类型</param>
        /// <returns> 转化后的指定类型的对象</returns>
        public static object CastTo(this object value, Type targetType)
        {
            if (value.IsNull())
            {
                return null;
            }
            if (targetType.IsNullableType())
            {
                targetType = targetType.GetUnderlyingType();
            }
            if (targetType.IsEnum)
            {
                return Enum.Parse(targetType, value.ToString());
            }
            if (targetType == typeof (Guid))
            {
                return Guid.Parse(value.ToString());
            }
            return Convert.ChangeType(value, targetType);
        }

        /// <summary>
        ///     把对象类型转化为指定类型
        /// </summary>
        /// <typeparam name="T"> 动态类型 </typeparam>
        /// <param name="value"> 要转化的源对象 </param>
        /// <returns> 转化后的指定类型的对象，转化失败引发异常。 </returns>
        public static T CastTo<T>(this object value)
        {
            object result = CastTo(value, typeof (T));
            return (T) result;
        }
    }
}