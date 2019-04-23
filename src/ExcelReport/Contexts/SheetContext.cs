using ExcelReport.Accumulations;
using ExcelReport.Driver;
using ExcelReport.Extends;
using ExcelReport.Meta;
using System;

namespace ExcelReport.Contexts
{
    public sealed class SheetContext
    {
        private readonly RowIndexAccumulation _rowIndexAccumulation = new RowIndexAccumulation();
        private readonly ISheet _sheet;

        public SheetContainer WorksheetContainer { get; }

        public bool IsEmpty()
        {
            return _sheet.IsNull();
        }

        public SheetContext(ISheet sheet, SheetContainer worksheetContainer)
        {
            _sheet = sheet;
            WorksheetContainer = worksheetContainer;
        }

        public ICell GetCell(Location location)
        {
            var rowIndex = _rowIndexAccumulation.GetCurrentRowIndex(location.RowIndex);

            IRow row = _sheet[rowIndex];
            if (row.IsNull())
            {
                return null;
            }
            return row[location.ColumnIndex];
        }

        public void CopyRepeaterTemplate(Repeater repeater, Action processTemplate)
        {
            var startRowIndex = _rowIndexAccumulation.GetCurrentRowIndex(repeater.Start.RowIndex);
            var endRowIndex = _rowIndexAccumulation.GetCurrentRowIndex(repeater.End.RowIndex);

            int span = _sheet.CopyRows(startRowIndex, endRowIndex);
            ICell startCell = GetCell(repeater.Start);
            startCell.Value = startCell.GetStringValue().CutEndOf($"<[{repeater.Name}]");
            processTemplate();
            ICell endCell = GetCell(repeater.End);
            endCell.Value = endCell.GetStringValue().CutStartOf($">[{repeater.Name}]");
            _rowIndexAccumulation.Add(span);
        }

        public void RemoveRepeaterTemplate(Repeater repeater)
        {
            var startRowIndex = _rowIndexAccumulation.GetCurrentRowIndex(repeater.Start.RowIndex);
            var endRowIndex = _rowIndexAccumulation.GetCurrentRowIndex(repeater.End.RowIndex);

            int span = _sheet.RemoveRows(startRowIndex, endRowIndex);
            _rowIndexAccumulation.Add(-span);
        }
    }
}