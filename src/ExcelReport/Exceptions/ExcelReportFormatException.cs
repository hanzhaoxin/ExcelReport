/*
 类：ExcelReportFormatException
 描述：ExcelReport格式化异常
 编 码 人：韩兆新 日期：2015年07月13日
 修改记录：

*/
namespace ExcelReport.Exceptions
{
    public class ExcelReportFormatException : ExcelReportException
    {
        public ExcelReportFormatException(string message) : base(message)
        {
        }
    }
}