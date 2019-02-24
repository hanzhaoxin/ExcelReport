using ExcelReport.Exceptions;
using ExcelReport.Extends;
using ExcelReport.Meta;
using NPOI.Extend;
using NPOI.SS.UserModel;
using System;

namespace ExcelReport.Renderers
{
    public class ParameterRenderer : Named, IElementRenderer
    {
        protected object Value { set; get; }

        public ParameterRenderer(string name, object value)
        {
            Name = name;
            Value = value;
        }

        public void Render(SheetAdapter sheetAdapter)
        {
            Parameter parameter = sheetAdapter.WorksheetContainer.Parameters[Name];
            foreach (var location in parameter.Locations)
            {
                ICell cell = sheetAdapter.GetCell(location);
                if (null == cell)
                {
                    throw new ExcelReportRenderException($"parameter[{parameter.Name}],cell[{location.RowIndex},{location.ColumnIndex}] is null");
                }
                var parameterName = $"$[{parameter.Name}]";
                if (parameterName.Equals(cell.StringCellValue.Trim()))
                {
                    cell.SetValue(Value);
                }
                else
                {
                    cell.SetValue(cell.StringCellValue.Replace(parameterName, Value.CastTo<string>()));
                }
            }
        }
    }

    public class ParameterRenderer<TSource> : Named, IEmbeddedRenderer<TSource>
    {
        protected Func<TSource, object> DgSetValue { set; get; }

        public ParameterRenderer(string name, Func<TSource, object> dgSetValue)
        {
            Name = name;
            DgSetValue = dgSetValue;
        }

        public void Render(SheetAdapter sheetAdapter, TSource dataSource)
        {
            Parameter parameter = sheetAdapter.WorksheetContainer.Parameters[Name];
            foreach (var location in parameter.Locations)
            {
                ICell cell = sheetAdapter.GetCell(location);
                if (null == cell)
                {
                    throw new ExcelReportRenderException($"parameter[{parameter.Name}],cell[{location.RowIndex},{location.ColumnIndex}] is null");
                }

                var parameterName = $"$[{parameter.Name}]";
                if (parameterName.Equals(cell.StringCellValue.Trim()))
                {
                    cell.SetValue(DgSetValue(dataSource));
                }
                else
                {
                    cell.SetValue(cell.StringCellValue.Replace($"$[{parameter.Name}]", DgSetValue(dataSource).CastTo<string>()));
                }
            }
        }
    }
}