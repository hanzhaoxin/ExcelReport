using ExcelReport.Meta;
using ExcelReport.Parsers;
using NPOI.SS.UserModel;

namespace ExcelReport
{
    public sealed class WorkbookAdapter
    {
        private static readonly TemplateParser TEMPLATE_PARSER = new TemplateParser();
        private readonly IWorkbook _workbook;

        private readonly WorkbookContainer _workbookContainer;

        public WorkbookAdapter(IWorkbook workbook)
        {
            _workbook = workbook;
            _workbookContainer = TEMPLATE_PARSER.Parse(workbook);
        }

        public SheetAdapter this[string sheetName]
        {
            get
            {
                var sheet = _workbook.GetSheet(sheetName);
                var worksheetContainer = _workbookContainer.Sheets[sheetName];
                return new SheetAdapter(sheet, worksheetContainer);
            }
        }
    }
}