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

        #region 属性

        protected int TemplateRowIndex
        {
            get { return _templateRowIndex; }
        }

        protected IEnumerable<TSource> DataSource
        {
            get { return _dataSource; }
        }

        protected List<TableColumnInfo<TSource>> ColumnInfoList
        {
            get { return _columnInfoList; }
        }

        #endregion

        /// 构造函数
        /// <param name="templateRowIndex"></param>
        /// <param name="dataSource"></param>
        /// <param name="columnInfos"></param>
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

        /// 格式化操作
        /// <param name="context"></param>
        public override void Format(SheetFormatterContext context)
        {
            context.ClearRowContent(TemplateRowIndex); //清除模板行单元格内容
            if (null == ColumnInfoList || ColumnInfoList.Count <= 0 || null == DataSource)
            {
                return;
            }
            var itemCount = 0;
            foreach (TSource rowSource in DataSource)
            {
                if (itemCount++ > 0)
                {
                    context.InsertEmptyRow(TemplateRowIndex);  //追加空行
                }
                var row = context.Sheet.GetRow(context.GetCurrentRowIndex(TemplateRowIndex));
                foreach (TableColumnInfo<TSource> colInfo in ColumnInfoList)
                {
                    var cell = row.GetCell(colInfo.ColumnIndex);
                    SetCellValue(cell, colInfo.DgSetValue(rowSource));
                }
            }
        }

        /// 添加列信息
        /// <param name="columnInfo"></param>
        public void AddColumnInfo(TableColumnInfo<TSource> columnInfo)
        {
            _columnInfoList.Add(columnInfo);
        }

        /// 添加列信息
        /// <param name="columnIndex"></param>
        /// <param name="dgSetValue"></param>
        public void AddColumnInfo(int columnIndex, Func<TSource, object> dgSetValue)
        {
            _columnInfoList.Add(new TableColumnInfo<TSource>(columnIndex, dgSetValue));
        }

    }
}