/*
 类：MergedRegionInfo
 描述：合并区域信息
 编 码 人：韩兆新 日期：2015年05月01日
 修改记录：

*/

namespace ExcelReport
{
    public class MergedRegionInfo
    {
        public MergedRegionInfo(int firstRow,int lastRow,int firstCol,int lastCol)
        {
            this.FirstRow = firstRow;
            this.LastRow = lastRow;
            this.FirstCol = firstCol;
            this.LastCol = lastCol;
        }
        public int FirstRow { get; set; }
        public int LastRow { get; set; }
        public int FirstCol { get; set; }
        public int LastCol { get; set; }
    }
}
