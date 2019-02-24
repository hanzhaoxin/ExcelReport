namespace ExcelReport.Accumulations
{
    public class RowIndexAccumulation : Accumulation
    {
        public int GetCurrentRowIndex(int sourceRowIndex)
        {
            return Value + sourceRowIndex;
        }
    }
}