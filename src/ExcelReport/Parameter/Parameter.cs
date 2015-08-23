/*
 类：Parameter
 描述：参数信息
 编 码 人：韩兆新 日期：2015年01月17日
 修改记录：
 * 2015年08月01日  删除CellPoint，添加RowIndex、ColumnIndex。

*/

namespace ExcelReport
{
    public class Parameter
    {
        public string Name { set; get; }
        public int RowIndex { set; get; }
        public int ColumnIndex { set; get; }
    }
}