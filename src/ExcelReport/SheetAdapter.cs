/*
 类：SheetAdapter
 描述：Sheet适配器
 编 码 人：韩兆新 日期：2015年01月17日
 修改记录：
 * 2015年08月02日（韩兆新）  修改自ExcelReport_v1.xx中的SheetFormatterContext。

*/
using System;
using NPOI.SS.UserModel;
using NPOI.Extend;

namespace ExcelReport
{
    public sealed class SheetAdapter
    {
        #region 成员字段

        private readonly RowIndexAccumulation _rowIndexAccumulation = new RowIndexAccumulation();
        private readonly ISheet _sheet; //当前格式化的Sheet

        #endregion

        #region 0 构造函数

        /// <summary>
        ///     构造函数
        /// </summary>
        /// <param name="sheet"></param>
        public SheetAdapter(ISheet sheet)
        {
            _sheet = sheet;
        }

        #endregion

        #region 1.0 获取行

        /// <summary>
        ///     根据行标获取行
        /// </summary>
        /// <param name="rowIndex">行标</param>
        /// <returns></returns>
        public IRow GetRow(int rowIndex)
        {
            rowIndex = _rowIndexAccumulation.GetCurrentRowIndex(rowIndex);

            return _sheet.GetRow(rowIndex);
        }

        #endregion

        #region 1.1 插入行

        /// <summary>
        ///     插入行
        /// </summary>
        /// <param name="rowIndex">行标</param>
        /// <returns>插入的行</returns>
        public IRow InsertRow(int rowIndex)
        {
            rowIndex = _rowIndexAccumulation.GetCurrentRowIndex(rowIndex);

            _rowIndexAccumulation.Add(1);
            return _sheet.InsertRow(rowIndex);
        }

        /// <summary>
        ///     插入行
        /// </summary>
        /// <param name="rowIndex">行标</param>
        /// <param name="insertRowsCount">插入的行数</param>
        /// <returns>插入的行</returns>
        public IRow[] InsertRows(int rowIndex, int rowsCount)
        {
            rowIndex = _rowIndexAccumulation.GetCurrentRowIndex(rowIndex);

            _rowIndexAccumulation.Add(rowsCount);
            return _sheet.InsertRows(rowIndex, rowsCount);
        }

        #endregion

        #region 1.2 删除行

        /// <summary>
        ///     删除行
        /// </summary>
        /// <param name="rowIndex">行标</param>
        public void RemoveRow(int rowIndex)
        {
            rowIndex = _rowIndexAccumulation.GetCurrentRowIndex(rowIndex);

            int span = _sheet.RemoveRow(rowIndex);
            _rowIndexAccumulation.Add(-span);
        }

        /// 删除指定行
        /// <param name="startRowIndex">开始行在模板中的行标</param>
        /// <param name="endRowIndex">结束行在模板中的行标</param>
        public void RemoveRows(int startRowIndex, int endRowIndex)
        {
            startRowIndex = _rowIndexAccumulation.GetCurrentRowIndex(startRowIndex);
            endRowIndex = _rowIndexAccumulation.GetCurrentRowIndex(endRowIndex);

            int span = _sheet.RemoveRows(startRowIndex, endRowIndex);
            _rowIndexAccumulation.Add(-span);
        }

        #endregion

        #region 1.3 复制行

        /// <summary>
        ///     复制行
        /// </summary>
        /// <param name="rowIndex">复制行行标</param>
        /// <returns>待追加的行标增量</returns>
        public void CopyRow(int rowIndex)
        {
            rowIndex = _rowIndexAccumulation.GetCurrentRowIndex(rowIndex);

            int span = _sheet.CopyRow(rowIndex);
            _rowIndexAccumulation.Add(span);
        }

        /// <summary>
        ///     复制行
        /// </summary>
        /// <param name="startRowIndex">开始行行标</param>
        /// <param name="endRowIndex">结束行行标</param>
        /// <returns>待追加的行标增量</returns>
        public void CopyRows(int startRowIndex, int endRowIndex)
        {
            startRowIndex = _rowIndexAccumulation.GetCurrentRowIndex(startRowIndex);
            endRowIndex = _rowIndexAccumulation.GetCurrentRowIndex(endRowIndex);

            int span = _sheet.CopyRows(startRowIndex, endRowIndex);
            _rowIndexAccumulation.Add(span);
        }

        #endregion

        #region 1.4 复制行并对源行做处理

        /// <summary>
        ///     复制行并对模板行做处理
        /// </summary>
        /// <param name="rowIndex">复制行行标</param>
        /// <param name="processTemplate">处理模板行的方法</param>
        public void CopyRow(int rowIndex, Action processTemplate)
        {
            rowIndex = _rowIndexAccumulation.GetCurrentRowIndex(rowIndex);

            int span = _sheet.CopyRow(rowIndex);
            processTemplate();
            _rowIndexAccumulation.Add(span);
        }

        /// <summary>
        ///     复制行并对模板行做处理
        /// </summary>
        /// <param name="startRowIndex">开始行行标</param>
        /// <param name="endRowIndex">结束行行标</param>
        /// <param name="processTemplate">处理模板行的方法</param>
        public void CopyRows(int startRowIndex, int endRowIndex, Action processTemplate)
        {
            startRowIndex = _rowIndexAccumulation.GetCurrentRowIndex(startRowIndex);
            endRowIndex = _rowIndexAccumulation.GetCurrentRowIndex(endRowIndex);

            int span = _sheet.CopyRows(startRowIndex, endRowIndex);
            processTemplate();
            _rowIndexAccumulation.Add(span);
        }

        #endregion
    }
}