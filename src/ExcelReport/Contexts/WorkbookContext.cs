using ExcelReport.Driver;
using ExcelReport.Meta;
using ExcelReport.Parsers;

namespace ExcelReport.Contexts
{
    public sealed class WorkbookContext
    {
        private static readonly TemplateParser TEMPLATE_PARSER = new TemplateParser();
        private readonly IWorkbook _workbook;

        private readonly WorkbookContainer _workbookContainer;

        public WorkbookContext(IWorkbook workbook)
        {
            _workbook = workbook;
            _workbookContainer = TEMPLATE_PARSER.Parse(workbook);
        }

        public SheetContext this[string sheetName]
        {
            get
            {
                var sheet = _workbook[sheetName];
                var worksheetContainer = _workbookContainer.Sheets[sheetName];
                return new SheetContext(sheet, worksheetContainer);
            }
        }
    }
}