/*
 类：ParseTemplate
 描述：解析模板
 编 码 人：韩兆新 日期：2015年01月17日
 修改记录：

*/

using System.Text.RegularExpressions;
using ExcelReport;
using NPOI.Extend;
using NPOI.SS.UserModel;

namespace XmlGenerator
{
    internal static class ParseTemplate
    {
        public static WorkbookParameterContainer Parse(string templatePath)
        {
            var workbookParameterContainer = new WorkbookParameterContainer();
            IWorkbook workbook = NPOIHelper.LoadWorkbook(templatePath);
            foreach (ISheet sheet in workbook)
            {
                workbookParameterContainer[sheet.SheetName] = new SheetParameterContainer
                {
                    SheetName = sheet.SheetName
                };
                foreach (IRow row in sheet)
                {
                    foreach (ICell cell in row.Cells)
                    {
                        if (cell.CellType.Equals(CellType.String))
                        {
                            MatchCollection matches = new Regex(@"(?<=\$\[)([\w]*)(?=\])").Matches(cell.StringCellValue);
                            foreach (Match match in matches)
                            {
                                var parameter = workbookParameterContainer[sheet.SheetName][match.Value];
                                if (parameter.IsNull())
                                {
                                    parameter = new Parameter();
                                    parameter.Name = match.Value;
                                    workbookParameterContainer[sheet.SheetName][match.Value] = parameter;
                                }
                                parameter.CellPointList.Add(new CellPoint {
                                    RowIndex = cell.RowIndex,
                                    ColumnIndex = cell.ColumnIndex
                                });
                            }
                        }
                    }
                }
            }
            return workbookParameterContainer;
        }
    }
}