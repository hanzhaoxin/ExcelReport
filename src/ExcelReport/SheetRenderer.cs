using ExcelReport.Extends;
using ExcelReport.Meta;
using ExcelReport.Renderers;
using System.Collections.Generic;

namespace ExcelReport
{
    public sealed class SheetRenderer : Named
    {
        private IList<IElementRenderer> RendererList { set; get; }

        public SheetRenderer(string sheetName, params IElementRenderer[] elementRenderers)
        {
            Name = sheetName;
            RendererList = new List<IElementRenderer>(elementRenderers);
        }

        public void Render(WorkbookAdapter workbookAdapter)
        {
            var worksheetAdapter = workbookAdapter[Name];
            if (worksheetAdapter.IsNull() || worksheetAdapter.IsEmpty())
            {
                return;
            }
            foreach (var renderer in RendererList)
            {
                renderer.Render(worksheetAdapter);
            }
        }
    }
}