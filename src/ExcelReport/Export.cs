using NPOI.Extend;
using NPOI.SS.UserModel;

namespace ExcelReport
{
    public sealed class Export
    {
        public static byte[] ExportToBuffer(string templateFile, params SheetRenderer[] sheetRenderers)
        {
            IWorkbook workbook = NPOIHelper.LoadWorkbook(templateFile);
            var workbookAdapter = new WorkbookAdapter(workbook);
            foreach (SheetRenderer sheetRenderer in sheetRenderers)
            {
                sheetRenderer.Render(workbookAdapter);
            }
            return workbook.SaveToBuffer();
        }
    }
}