/*
 类：PictureStyle
 描述：图片样式
 编 码 人：韩兆新 日期：2015年04月11日
 修改记录：

*/
using NPOI.SS.UserModel;

namespace ExcelReport
{
    public class PictureStyle
    {
        public int AnchorDx1 { get; set; }
        public int AnchorDx2 { get; set; }
        public int AnchorDy1 { get; set; }
        public int AnchorDy2 { get; set; }

        public int FillColor { get; set; }
        public bool IsNoFill { get; set; }
        public LineStyle LineStyle { get; set; }
        public int LineStyleColor { get; set; }
        public double LineWidth { get; set; }
    }
}
