/*
 类：TypeExtend
 描述：Type的扩展方法
 编 码 人：韩兆新 日期：2015年08月01日
 修改记录：

*/
using System;
using System.ComponentModel;

namespace ExcelReport
{
    public static class TypeExtend
    {
        /// <summary>
        ///     判断类型是否为Nullable类型
        /// </summary>
        /// <param name="type"> 要处理的类型 </param>
        /// <returns> 是返回True，不是返回False </returns>
        public static bool IsNullableType(this Type type)
        {
            return ((type != null) && type.IsGenericType) && (type.GetGenericTypeDefinition() == typeof (Nullable<>));
        }

        /// <summary>
        ///     通过类型转换器获取Nullable类型的基础类型
        /// </summary>
        /// <param name="type"> 要处理的类型对象 </param>
        /// <returns> 对应的基础类型</returns>
        public static Type GetUnderlyingType(this Type type)
        {
            if (IsNullableType(type))
            {
                var nullableConverter = new NullableConverter(type);
                return nullableConverter.UnderlyingType;
            }
            return type;
        }
    }
}