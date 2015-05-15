/*
 类：SheetFormatterContext
 描述：Sheet格式化的上下文
 编 码 人：韩兆新 日期：2015年01月17日
 修改记录：
    1: 修改人：韩兆新  日期：2015年04月11日
       修改内容：添加了复制行组的方法CopyRows();
                 添加了删除行组的方法RemoveRows();
    2: 修改人：韩兆新  日期：2015年05月01日
       修改内容：修改了复制行组的方法CopyRows(),解决了跨多行的合并单元格复制丢失合并信息的问题;
*/

using System.Collections.Generic;
using NPOI.SS.UserModel;

namespace ExcelReport
{
    public class SheetFormatterContext
    {
        #region 成员字段及属性

        private int _increaseRowsCount = 0;

        public ISheet Sheet { get; set; }

        public IEnumerable<ElementFormatter> Formatters { get; set; }

        #endregion 成员字段及属性

        /// 构造函数
        /// 
        public SheetFormatterContext()
        {
        }

        /// 构造函数
        /// <param name="sheet"></param>
        /// <param name="formatters"></param>
        public SheetFormatterContext(ISheet sheet, IEnumerable<ElementFormatter> formatters)
        {
            this.Sheet = sheet;
            this.Formatters = formatters;
        }

        
        /// 获取指定行当前行标
        /// <param name="rowIndexInTemplate">指定行在模板中的行标</param>
        /// <returns>当前行标</returns>
        public int GetCurrentRowIndex(int rowIndexInTemplate)
        {
            return rowIndexInTemplate + _increaseRowsCount;
        }

       
        /// 在指定行后插入一行（并将指定行作为模板复制样式）
        /// <param name="templateRowIndex">模板行在模板中的行标</param>
        public void InsertEmptyRow(int templateRowIndex)
        {
            var templateRow = Sheet.GetRow(GetCurrentRowIndex(templateRowIndex));
            var insertRowIndex = GetCurrentRowIndex(templateRowIndex + 1);
            if (insertRowIndex < Sheet.LastRowNum)
            {
                Sheet.ShiftRows(insertRowIndex, Sheet.LastRowNum, 1, true, false);
            }
            var newRow = Sheet.CreateRow(GetCurrentRowIndex(templateRowIndex + 1));
            newRow.Height = templateRow.Height;
            foreach (var cell in templateRow.Cells)
            {
                newRow.CreateCell(cell.ColumnIndex).CellStyle = cell.CellStyle;
            }
            //获取模板行内的合并区域
            var regionInfoList = Sheet.GetAllMergedRegionInfos(GetCurrentRowIndex(templateRowIndex), GetCurrentRowIndex(templateRowIndex), null, null);
            //复制合并区域
            foreach (var regionInfo in regionInfoList)
            {
                regionInfo.FirstRow += 1;
                regionInfo.LastRow += 1;
                Sheet.AddMergedRegion(regionInfo);
            }
            _increaseRowsCount++;
        }

        
        /// 清除指定行单元格内容
        /// <param name="rowIndex">指定行在模板中的行标</param>
        public void ClearRowContent(int rowIndex)
        {
            var row = Sheet.GetRow(GetCurrentRowIndex(rowIndex));
            foreach (var cell in row.Cells)
            {
                cell.SetCellValue(string.Empty);
            }
        }

        /// 删除指定行
        /// <param name="rowIndex">指定行在模板中的行标</param>
        public void RemoveRow(int rowIndex)
        {
            var row = Sheet.GetRow(GetCurrentRowIndex(rowIndex));
            Sheet.RemoveRow(row);
        }
        /// 删除指定行
        /// <param name="startRowIndex">开始行在模板中的行标</param>
        /// <param name="endRowIndex">结束行在模板中的行标</param>
        public void RemoveRows(int startRowIndex,int endRowIndex)
        {
            for (int i = startRowIndex; i <= endRowIndex; i++)
            {
                RemoveRow(i);
            }
        }

        /// 复制模板行
        /// <param name="templateStartRowIndex"></param>
        /// <param name="templateEndRowIndex"></param>
        public void CopyRows(int templateStartRowIndex, int templateEndRowIndex)
        {
            var span = templateEndRowIndex - templateStartRowIndex + 1;
            var insertStartRowIndex = GetCurrentRowIndex(templateStartRowIndex + span);
            if (insertStartRowIndex < Sheet.LastRowNum)
            {
                Sheet.ShiftRows(insertStartRowIndex, Sheet.LastRowNum, span, true, false);
            }
            for (int i = templateStartRowIndex; i <= templateEndRowIndex; i++)
            {
                var row = Sheet.GetRow(GetCurrentRowIndex(i));
                var newRow = row.CopyTo(GetCurrentRowIndex(i + span));
            }
            //获取模板行内的合并区域
            var regionInfoList = Sheet.GetAllMergedRegionInfos(GetCurrentRowIndex(templateStartRowIndex), GetCurrentRowIndex(templateEndRowIndex), null, null);
            //复制合并区域
            foreach (var regionInfo in regionInfoList)
            {
                regionInfo.FirstRow += span;
                regionInfo.LastRow += span;
                Sheet.AddMergedRegion(regionInfo);
            }
            //获取模板行内的图片
            var picInfoList = Sheet.GetAllPictureInfos(GetCurrentRowIndex(templateStartRowIndex), GetCurrentRowIndex(templateEndRowIndex), null, null);
            //复制图片
            foreach (var picInfo in picInfoList)
            {
                picInfo.MaxRow += span;
                picInfo.MinRow += span;
                Sheet.AddPicture(picInfo);
            }
            _increaseRowsCount += span;
        }

        /// 格式化Sheet
        /// 
        public void Format()
        {
            if (null == Sheet || null == Formatters)
            {
                return;
            }
            foreach (var formatter in Formatters)
            {
                formatter.Format(this);
            }
        }
    }
}