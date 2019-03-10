using System;
using System.Text;

namespace ExcelReport.Driver.CSV.Extends
{
    internal static class StringExtend
    {
        public static bool ContainAny(this string self, params char[] chars)
        {
            return self.IndexOfAny(chars) >= 0;
        }

        public static byte[] ToBytesBy<TEncoding>(this string str)
            where TEncoding : Encoding
        {
            var enc = Activator.CreateInstance<TEncoding>();
            return enc.GetBytes(str);
        }

        public static byte[] ToBytesBy(this string str, Encoding encoding)
        {
            return encoding.GetBytes(str);
        }
    }
}