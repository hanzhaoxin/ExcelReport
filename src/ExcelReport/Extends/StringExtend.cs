namespace ExcelReport.Extends
{
    public static class StringExtend
    {
        public static string CutStartOf(this string self, string value)
        {
            var idx = self.IndexOf(value);
            if (idx < 0)
            {
                return self;
            }
            return self.Remove(idx);
        }

        public static string CutEndOf(this string self, string value)
        {
            var idx = self.LastIndexOf(value);
            if (idx < 0)
            {
                return self;
            }
            return self.Remove(0, idx + value.Length);
        }
    }
}