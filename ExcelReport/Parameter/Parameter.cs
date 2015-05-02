/*
 类：Parameter
 描述：参数信息
 编 码 人：韩兆新 日期：2015年01月17日
 修改记录：

*/

using System.Drawing;

namespace ExcelReport
{
    public class Parameter
    {
        public Parameter()
        {
        }

        public Parameter(string sheetName, string parameterName, Point cellPoint)
        {
            this.SheetName = sheetName;
            this.ParameterName = parameterName;
            this.CellPoint = cellPoint;
        }

        public string SheetName { set; get; }

        public string ParameterName { set; get; }

        public Point CellPoint { set; get; }
    }
}