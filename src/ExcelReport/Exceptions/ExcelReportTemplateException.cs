/*
 类：ExcelReportTemplateException
 描述：ExcelReport模板异常
 编 码 人：韩兆新 日期：2015年07月13日
 修改记录：

*/
namespace ExcelReport.Exceptions
{
    internal class ExcelReportTemplateException : ExcelReportException
    {
        public ExcelReportTemplateException(string message)
            : base(message)
        {
        }
    }
}