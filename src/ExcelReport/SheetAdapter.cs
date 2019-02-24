using ExcelReport.Accumulations;
using ExcelReport.Extends;
using ExcelReport.Meta;
using NPOI.Extend;
using NPOI.SS.UserModel;
using System;

namespace ExcelReport
{
    public sealed class SheetAdapter
    {
        private readonly RowIndexAccumulation _rowIndexAccumulation = new RowIndexAccumulation();
        private readonly ISheet _sheet;

        public SheetContainer WorksheetContainer { get; }

        public bool IsEmpty()
        {
            return _sheet.IsNull();
        }

        public SheetAdapter(ISheet sheet, SheetContainer worksheetContainer)
        {
            _sheet = sheet;
            WorksheetContainer = worksheetContainer;
        }

        public ICell GetCell(Location location)
        {
            var rowIndex = _rowIndexAccumulation.GetCurrentRowIndex(location.RowIndex);

            IRow row = _sheet.GetRow(rowIndex);
            if (row.IsNull())
            {
                return null;
            }
            return row.GetCell(location.ColumnIndex);
        }

        public void CopyRepeaterTemplate(Repeater repeater, Action processTemplate)
        {
            var startRowIndex = _rowIndexAccumulation.GetCurrentRowIndex(repeater.Start.RowIndex);
            var endRowIndex = _rowIndexAccumulation.GetCurrentRowIndex(repeater.End.RowIndex);

            int span = _sheet.CopyRows(startRowIndex, endRowIndex);
            ICell startCell = GetCell(repeater.Start);
            startCell.SetValue(startCell.StringCellValue.Replace($"<[{repeater.Name}]", String.Empty));
            processTemplate();
            ICell endCell = GetCell(repeater.End);
            endCell.SetValue(endCell.StringCellValue.Replace($">[{repeater.Name}]", String.Empty));
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