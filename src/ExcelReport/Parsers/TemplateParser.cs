using ExcelReport.Driver;
using ExcelReport.Extends;
using ExcelReport.Meta;

namespace ExcelReport.Parsers
{
    public sealed class TemplateParser
    {
        private static readonly ParameterParser PARAMETER_PARSER = new ParameterParser();

        private static readonly RepeaterStartParser REPEATER_START_PARSER = new RepeaterStartParser();

        private static readonly RepeaterEndParser REPEATER_END_PARSER = new RepeaterEndParser();

        public WorkbookContainer Parse(IWorkbook workbook)
        {
            WorkbookContainer workbookContainer = new WorkbookContainer();
            foreach (ISheet sheet in workbook)
            {
                foreach (IRow row in sheet)
                {
                    foreach (ICell cell in row)
                    {
                        if (cell.Value is string)
                        {
                            foreach (var parameterName in PARAMETER_PARSER.Parse(cell.GetStringValue()))
                            {
                                workbookContainer.Sheets[sheet.SheetName].Parameters[parameterName].Append(new Location(cell.RowIndex, cell.ColumnIndex));
                            }

                            foreach (var tagName in REPEATER_START_PARSER.Parse(cell.GetStringValue()))
                            {
                                workbookContainer.Sheets[sheet.SheetName].Repeaters[tagName].Start = new Location(cell.RowIndex, cell.ColumnIndex);
                            }

                            foreach (var tagName in REPEATER_END_PARSER.Parse(cell.GetStringValue()))
                            {
                                workbookContainer.Sheets[sheet.SheetName].Repeaters[tagName].End = new Location(cell.RowIndex, cell.ColumnIndex);
                            }
                        }
                    }
                }
            }

            return workbookContainer;
        }
    }
}