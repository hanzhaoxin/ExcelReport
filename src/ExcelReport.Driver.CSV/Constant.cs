namespace ExcelReport.Driver.CSV
{
    internal static class Constant
    {
        public const char ESCAPE = '\"';

        public const char DELIMITER = ',';

        public const string ROW_END = "\r\n";

        public static readonly char[] NEED_ESCAPE_CHARS = new char[4] { ESCAPE, DELIMITER, '\r', '\n' };
    }
}