namespace ExcelReport.Driver
{
    public interface ICell : IAdapter
    {
        int RowIndex { get; }

        int ColumnIndex { get; }

        object Value { get; set; }
    }
}