/*
 类：ExcelReportException
 描述：ExcelReport异常
 编 码 人：韩兆新 日期：2015年07月13日
 修改记录：

*/
using System;

namespace ExcelReport.Exceptions
{
    public class ExcelReportException : ApplicationException
    {
        public ExcelReportException(string message) : base(message)
        {
        }
    }
}