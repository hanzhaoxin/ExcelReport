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
