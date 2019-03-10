using System.Text;

namespace ExcelReport.Driver.CSV.Extends
{
    internal static class EncodingExtend
    {
        static EncodingExtend()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        public static Encoding GB2312
        {
            get
            {
                return Encoding.GetEncoding("GB2312");
            }
        }
    }
}