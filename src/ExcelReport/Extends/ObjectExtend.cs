using System;

namespace ExcelReport.Extends
{
    public static class ObjectExtend
    {
        public static bool IsNull(this object value)
        {
            return null == value;
        }

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
            if (targetType == typeof(Guid))
            {
                return Guid.Parse(value.ToString());
            }
            return Convert.ChangeType(value, targetType);
        }

        public static T CastTo<T>(this object value)
        {
            object result = CastTo(value, typeof(T));
            return (T)result;
        }
    }
}