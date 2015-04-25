using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.SS.UserModel;

namespace ExcelReport
{
    public class PicturesStyle
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
