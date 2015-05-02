using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace _01_生成工资表_工资条报表
{
    public class SalaryInfo
    {
        public string JobNo { get; set; }
        public string Name{ get; set; }
        public double HourlyWage{ get; set; }
        public int ManHour{ get; set; }
        public int UsualOvertime{ get; set; }
        public int WeekendOvertime{ get; set; }
        public int LegalOvertime{ get; set; }
        public double BasePay{ get; set; }
        public double OvertimePay{ get; set; }
        public double FoodAllowance{ get; set; }
        public double OtherAllowance{ get; set; }
        public double GrossPay{ get; set; }
        public double BoardWages{ get; set; }
        public double Utilities{ get; set; }
        public double BadDebt{ get; set; }
        public double RealWages{ get; set; }
        public string Remark { get; set; }
    }
}
