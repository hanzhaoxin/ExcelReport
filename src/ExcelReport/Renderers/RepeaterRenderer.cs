using ExcelReport.Exceptions;
using ExcelReport.Extends;
using ExcelReport.Meta;
using System;
using System.Collections.Generic;

namespace ExcelReport.Renderers
{
    public class RepeaterRenderer<TItem> : Named, IElementRenderer
    {
        protected IEnumerable<TItem> DataSource { set; get; }

        protected IList<IEmbeddedRenderer<TItem>> RendererList { set; get; }

        public RepeaterRenderer(string name, IEnumerable<TItem> dataSource, params IEmbeddedRenderer<TItem>[] renderers)
        {
            Name = name;
            DataSource = dataSource;
            RendererList = new List<IEmbeddedRenderer<TItem>>(renderers);
        }

        public void Render(SheetAdapter sheetAdapter)
        {
            Repeater repeater = sheetAdapter.WorksheetContainer.Repeaters[Name];
            if (RendererList.IsNullOrEmpty())
            {
                throw new ExcelReportRenderException($"RepeaterRenderer[{repeater.Name}] is empty");
            }

            foreach (var item in DataSource)
            {
                sheetAdapter.CopyRepeaterTemplate(repeater, () =>
                {
                    foreach (var renderer in RendererList)
                    {
                        renderer.Render(sheetAdapter, item);
                    }
                });
            }
            sheetAdapter.RemoveRepeaterTemplate(repeater);
        }

        public void Append(IEmbeddedRenderer<TItem> renderer)
        {
            RendererList.Add(renderer);
        }
    }

    public class RepeaterRenderer<TSource, TItem> : Named, IEmbeddedRenderer<TSource>
    {
        protected Func<TSource, IEnumerable<TItem>> DgSetDataSource { set; get; }

        protected IList<IEmbeddedRenderer<TItem>> RendererList { set; get; }

        public RepeaterRenderer(string name, Func<TSource, IEnumerable<TItem>> dgSetDataSource, params IEmbeddedRenderer<TItem>[] renderers)
        {
            Name = name;
            DgSetDataSource = dgSetDataSource;
            RendererList = new List<IEmbeddedRenderer<TItem>>(renderers);
        }

        public void Render(SheetAdapter sheetAdapter, TSource dataSource)
        {
            Repeater repeater = sheetAdapter.WorksheetContainer.Repeaters[Name];
            if (RendererList.IsNullOrEmpty())
            {
                throw new ExcelReportRenderException($"RepeaterRenderer[{repeater.Name}] is empty");
            }

            foreach (var item in DgSetDataSource(dataSource))
            {
                sheetAdapter.CopyRepeaterTemplate(repeater, () =>
                {
                    foreach (var renderer in RendererList)
                    {
                        renderer.Render(sheetAdapter, item);
                    }
                });
            }
            sheetAdapter.RemoveRepeaterTemplate(repeater);
        }

        public void Append(IEmbeddedRenderer<TItem> renderer)
        {
            RendererList.Add(renderer);
        }
    }
}