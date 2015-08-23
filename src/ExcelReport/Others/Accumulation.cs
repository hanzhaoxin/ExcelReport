/*
 类：Accumulation
 描述：累加器
 编 码 人：韩兆新 日期：2015年08月01日
 修改记录：
 * 

*/
namespace ExcelReport
{
    public class Accumulation
    {
        #region 成员字段及属性

        private int _value;

        public int Value
        {
            get { return _value; }
        }

        #endregion

        #region 0 构造函数
        /// <summary>
        /// 实例化累加器
        /// </summary>
        public Accumulation()
        {
            _value = 0;
        }

        #endregion

        #region 1.0 累加增量

        /// <summary>
        ///     累加增量
        /// </summary>
        /// <param name="increment">增量</param>
        public void Add(int increment)
        {
            _value += increment;
        }

        #endregion
    }
}