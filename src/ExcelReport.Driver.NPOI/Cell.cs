using NPOI.Extend;
using NpoiCell = NPOI.SS.UserModel.ICell;

namespace ExcelReport.Driver.NPOI
{
    public class Cell : ICell
    {
        public NpoiCell NpoiCell { get; private set; }

        public Cell(NpoiCell npoiCell)
        {
            NpoiCell = npoiCell;
        }

        public int RowIndex => NpoiCell.RowIndex;

        public int ColumnIndex => NpoiCell.ColumnIndex;

        public object Value { get => NpoiCell.GetValue(); set => NpoiCell.SetValue(value); }

        public object GetOriginal()
        {
            return NpoiCell;
        }
    }
}