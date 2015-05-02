/*
 结构：CellSpan
 描述：单元格跨度信息
 编 码 人：韩兆新 日期：2015年01月17日
 修改记录：

*/
namespace ExcelReport
{
    struct CellSpan
    {
        public CellSpan(int rowSpan, int colSpan)
        {
            this.rowSpan = rowSpan;
            this.colSpan = colSpan;
        }
        private int rowSpan;

        public int RowSpan
        {
            get { return rowSpan; }
            set { rowSpan = value; }
        }
        private int colSpan;

        public int ColSpan
        {
            get { return colSpan; }
            set { colSpan = value; }
        }
    }
}
