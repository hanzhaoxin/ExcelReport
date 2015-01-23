/*
 类：ElementFormatter
 描述：（元素）格式化器（抽象）
 编 码 人：韩兆新 日期：2015年01月17日
 修改记录：
    日期：2015年01月23日  修改人：韩兆新   修改内容描述：修改了SetCellValue()方法，添加了对图片的处理
 
*/

using System;
using NPOI.SS.UserModel;

namespace ExcelReport
{
    public abstract class ElementFormatter
    {
        #region 设置单元格值

        protected virtual void SetCellValue(ICell cell, object value)
        {
            if (null == cell)
            {
                return;
            }
            if (null == value)
            {
                cell.SetCellValue(string.Empty);
            }
            else
            {
                if (value.GetType().FullName.Equals("System.Byte[]"))
                {
                    var pictureIdx = cell.Sheet.Workbook.AddPicture((Byte[]) value, PictureType.PNG);
                    var anchor = cell.Sheet.Workbook.GetCreationHelper().CreateClientAnchor();
                    anchor.Col1 = cell.ColumnIndex;
                    anchor.Col2 = cell.ColumnIndex + cell.GetSpan().ColSpan;
                    anchor.Row1 = cell.RowIndex;
                    anchor.Row2 = cell.RowIndex + cell.GetSpan().RowSpan;

                    var patriarch = cell.Sheet.CreateDrawingPatriarch();
                    var pic = patriarch.CreatePicture(anchor, pictureIdx);
                }
                else
                {
                    var valueTypeCode = Type.GetTypeCode(value.GetType());
                    switch (valueTypeCode)
                    {
                        case TypeCode.String: //字符串类型
                            cell.SetCellValue(Convert.ToString(value));
                            break;

                        case TypeCode.DateTime: //日期类型
                            cell.SetCellValue(Convert.ToDateTime(value));
                            break;

                        case TypeCode.Boolean: //布尔型
                            cell.SetCellValue(Convert.ToBoolean(value));
                            break;

                        case TypeCode.Int16: //整型
                        case TypeCode.Int32:
                        case TypeCode.Int64:
                        case TypeCode.Byte:
                        case TypeCode.Single: //浮点型
                        case TypeCode.Double:
                        case TypeCode.UInt16: //无符号整型
                        case TypeCode.UInt32:
                        case TypeCode.UInt64:
                            cell.SetCellValue(Convert.ToDouble(value));
                            break;

                        default:
                            cell.SetCellValue(string.Empty);
                            break;
                    }
                }
            }
        }

        #endregion 设置单元格值

        #region 格式化操作

        public abstract void Format(SheetFormatterContext context);

        #endregion 格式化操作
    }
}