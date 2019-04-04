using ExcelReport.Contexts;
using ExcelReport.Extends;
using ExcelReport.Meta;
using System.Collections.Generic;
using System.Linq;

namespace ExcelReport.Renderers
{
    public class SheetRenderer : Named
    {
        private IList<IElementRenderer> RendererList { set; get; }

        public SheetRenderer(string sheetName, params IElementRenderer[] elementRenderers)
        {
            Name = sheetName;
            RendererList = new List<IElementRenderer>(elementRenderers);
        }

        public virtual void Render(WorkbookContext workbookContext)
        {
            var worksheetContext = workbookContext[Name];
            if (worksheetContext.IsNull() || worksheetContext.IsEmpty())
            {
                return;
            }
            foreach (var renderer in RendererList.OrderBy(renderer => renderer.SortNum(worksheetContext)))
            {
                renderer.Render(worksheetContext);
            }
        }
    }
}