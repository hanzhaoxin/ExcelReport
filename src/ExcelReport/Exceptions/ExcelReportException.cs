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