namespace ExcelReport.Exceptions
{
    internal class UnconfiguredException : ExcelReportException
    {
        public UnconfiguredException(string message) : base(message)
        {
        }
    }
}