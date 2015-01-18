/*
 类：TableFormatter
 描述：表格（元素）格式化器
 编 码 人：韩兆新 日期：2015年01月17日
 修改记录：

*/

using System;
using System.Collections.Generic;

namespace ExcelReport
{
    public class TableFormatter<TSource> : ElementFormatter
    {
        #region 成员字段

        private int _templateRowIndex;
        private IEnumerable<TSource> _dataSource;
        private List<TableColumnInfo<TSource>> _columnInfoList;

        #endregion 成员字段

        #region 构造函数

        public TableFormatter(int templateRowIndex, IEnumerable<TSource> dataSource, params TableColumnInfo<TSource>[] columnInfos)
        {
            _templateRowIndex = templateRowIndex;
            _dataSource = dataSource;
            _columnInfoList = new List<TableColumnInfo<TSource>>();
            if (null != columnInfos && columnInfos.Length > 0)
            {
                _columnInfoList.AddRange(columnInfos);
            }
        }

        #endregion 构造函数

        #region 格式化操作

        public override void Format(SheetFormatterContext context)
        {
            context.ClearRowContent(_templateRowIndex); //清除模板行单元格内容
            if (null == _columnInfoList || _columnInfoList.Count <= 0 || null == _dataSource)
            {
                return;
            }
            foreach (TSource rowSource in _dataSource)
            {
                var row = context.Sheet.GetRow(context.GetCurrentRowIndex(_templateRowIndex));
                foreach (TableColumnInfo<TSource> colInfo in _columnInfoList)
                {
                    var cell = row.GetCell(colInfo.ColumnIndex);
                    SetCellValue(cell, colInfo.DgSetValue(rowSource));
                }
                context.InsertEmptyRow(_templateRowIndex);  //追加空行
            }
            context.RemoveRow(_templateRowIndex);   //删除空行
        }

        #endregion 格式化操作

        #region 添加列信息

        public void AddColumnInfo(TableColumnInfo<TSource> columnInfo)
        {
            _columnInfoList.Add(columnInfo);
        }

        public void AddColumnInfo(int columnIndex, Func<TSource, object> dgSetValue)
        {
            _columnInfoList.Add(new TableColumnInfo<TSource>(columnIndex, dgSetValue));
        }

        #endregion 添加列信息
    }
}