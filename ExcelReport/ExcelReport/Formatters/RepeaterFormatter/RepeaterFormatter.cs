/*
 类：RepeaterFormatter
 描述：Repeater（元素）格式化器
 编 码 人：韩兆新 日期：2015年04月11日
 修改记录：

*/

using System;
using System.Collections.Generic;
using System.Drawing;
namespace ExcelReport
{
    public class RepeaterFormatter<TSource> : ElementFormatter
    {
        #region 成员字段
        private Point _startTagCell;
        private Point _endTagCell;
        private IEnumerable<TSource> _dataSource;
        private List<RepeaterCellInfo<TSource>> _cellInfoList;
        #endregion

        #region 属性
        protected Point StartTagCell
        {
            get { return _startTagCell; }
        }

        protected Point EndTagCell
        {
            get { return _endTagCell; }
        }

        protected IEnumerable<TSource> DataSource
        {
            get { return _dataSource; }
        }

        protected List<RepeaterCellInfo<TSource>> CellInfoList
        {
            get { return _cellInfoList; }
        }
        #endregion

        /// 构造函数
        /// <param name="startTagCell"></param>
        /// <param name="endTagCell"></param>
        /// <param name="dataSource"></param>
        /// <param name="cellInfos"></param>
        public RepeaterFormatter(Point startTagCell, Point endTagCell, IEnumerable<TSource> dataSource, params RepeaterCellInfo<TSource>[] cellInfos)
        {
            _startTagCell = startTagCell;
            _endTagCell = endTagCell;
            _dataSource = dataSource;
            _cellInfoList = new List<RepeaterCellInfo<TSource>>();
            if (null != cellInfos && cellInfos.Length > 0)
            {
                _cellInfoList.AddRange(cellInfos);
            }
        }

        /// 格式化操作
        /// <param name="context"></param>
        public override void Format(SheetFormatterContext context)
        {
            context.ClearRowContent(StartTagCell.X);
            context.ClearRowContent(EndTagCell.X);
            if (null == CellInfoList || CellInfoList.Count <= 0 || null == DataSource)
            {
                return;
            }
            var itemCount = 0;
            foreach (TSource itemSource in DataSource)
            {
                if (itemCount++ > 0)
                {
                    context.CopyRows(StartTagCell.X, EndTagCell.X);  //追加空行
                }
                foreach (RepeaterCellInfo<TSource> cellInfo in CellInfoList)
                {
                    var rowIndex = context.GetCurrentRowIndex(cellInfo.CellPoint.X);
                    var row = context.Sheet.GetRow(rowIndex) ?? context.Sheet.CreateRow(rowIndex);
                    var cell = row.GetCell(cellInfo.CellPoint.Y) ?? row.CreateCell(cellInfo.CellPoint.Y);
                    SetCellValue(cell, cellInfo.DgSetValue(itemSource));
                }
            }
        }

        /// 添加重复项单元格信息
        /// <param name="cellInfo"></param>
        public void AddCellInfo(RepeaterCellInfo<TSource> cellInfo)
        {
            _cellInfoList.Add(cellInfo);
        }

        /// 添加重复项单元格信息
        /// <param name="cellPoint"></param>
        /// <param name="dgSetValue"></param>
        public void AddCellInfo(Point cellPoint, Func<TSource, object> dgSetValue)
        {
            _cellInfoList.Add(new RepeaterCellInfo<TSource>(cellPoint, dgSetValue));
        }
    }
}
