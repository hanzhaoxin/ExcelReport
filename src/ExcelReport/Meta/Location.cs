namespace ExcelReport.Meta
{
    public class Location
    {
        public int RowIndex { private set; get; }
        public int ColumnIndex { private set; get; }

        public Location(int rowIndex, int columnIndex)
        {
            RowIndex = rowIndex;
            ColumnIndex = columnIndex;
        }
    }
}