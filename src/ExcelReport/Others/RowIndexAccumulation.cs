/*
 类：RowIndexAccumulation
 描述：行标累加器
 编 码 人：韩兆新 日期：2015年08月01日
 修改记录：
 * 

*/
namespace ExcelReport
{
    public class RowIndexAccumulation : Accumulation
    {
        #region 1.0 获取当前行标

        /// <summary>
        ///     获取当前行标
        /// </summary>
        /// <param name="sourceRowIndex">源行标</param>
        /// <returns></returns>
        public int GetCurrentRowIndex(int sourceRowIndex)
        {
            return Value + sourceRowIndex;
        }

        #endregion
    }
}