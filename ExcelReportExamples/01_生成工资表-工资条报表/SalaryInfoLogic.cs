using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace _01_生成工资表_工资条报表
{
    public static class SalaryInfoLogic
    {
        public static List<SalaryInfo> GetList()
        {
            List<SalaryInfo> list = new List<SalaryInfo>();

            list.Add(new SalaryInfo
            {
                JobNo = "006",
                Name = "梁忠基",
                HourlyWage = 11.31,
                ManHour = 174,
                UsualOvertime = 28,
                WeekendOvertime = 8,
                LegalOvertime = 0,
                BasePay = 1967.94,
                OvertimePay = 491.02,
                FoodAllowance= 0,
                OtherAllowance = 0,
                GrossPay = 2458.96,
                BoardWages = 60,
                Utilities = 30,
                BadDebt = 90,
                RealWages = 2368.96,
                Remark = string.Empty
            });

            list.Add(new SalaryInfo
            {
                JobNo = "007",
                Name = "韦明珍",
                HourlyWage = 6.68,
                ManHour = 174,
                UsualOvertime = 28,
                WeekendOvertime = 8,
                LegalOvertime = 0,
                BasePay = 1162.32,
                OvertimePay = 296.56,
                FoodAllowance = 0,
                OtherAllowance = 0,
                GrossPay = 1458.88,
                BoardWages = 60,
                Utilities = 30,
                BadDebt = 90,
                RealWages = 1368.88,
                Remark = string.Empty
            });

            list.Add(new SalaryInfo
            {
                JobNo = "008",
                Name = "韦明初",
                HourlyWage = 5.46,
                ManHour = 174,
                UsualOvertime = 28,
                WeekendOvertime = 8,
                LegalOvertime = 0,
                BasePay = 950.04,
                OvertimePay = 245.32,
                FoodAllowance = 0,
                OtherAllowance = 0,
                GrossPay = 1195.36,
                BoardWages = 60,
                Utilities = 30,
                BadDebt = 90,
                RealWages = 1105.36,
                Remark = string.Empty
            });

            list.Add(new SalaryInfo
            {
                JobNo = "010",
                Name = "钟燕华",
                HourlyWage = 5.25,
                ManHour = 174,
                UsualOvertime = 28,
                WeekendOvertime = 8,
                LegalOvertime = 0,
                BasePay = 913.5,
                OvertimePay = 236.5,
                FoodAllowance = 0,
                OtherAllowance = 0,
                GrossPay = 1150,
                BoardWages = 60,
                Utilities = 30,
                BadDebt = 90,
                RealWages = 1060,
                Remark = string.Empty
            });

            list.Add(new SalaryInfo
            {
                JobNo = "011",
                Name = "谭兴燕",
                HourlyWage = 7.19,
                ManHour = 174,
                UsualOvertime = 28,
                WeekendOvertime = 8,
                LegalOvertime = 0,
                BasePay = 1251.06,
                OvertimePay = 317.98,
                FoodAllowance = 0,
                OtherAllowance = 0,
                GrossPay = 1569.04,
                BoardWages = 60,
                Utilities = 30,
                BadDebt = 90,
                RealWages = 1479.04,
                Remark = string.Empty
            });

            list.Add(new SalaryInfo
            {
                JobNo = "012",
                Name = "梁海玲",
                HourlyWage = 6.76,
                ManHour = 174,
                UsualOvertime = 28,
                WeekendOvertime = 8,
                LegalOvertime = 0,
                BasePay = 1176.24,
                OvertimePay = 299.92,
                FoodAllowance = 0,
                OtherAllowance = 0,
                GrossPay = 1476.16,
                BoardWages = 60,
                Utilities = 30,
                BadDebt = 90,
                RealWages = 1386.16,
                Remark = string.Empty
            });

            list.Add(new SalaryInfo
            {
                JobNo = "013",
                Name = "黄金红",
                HourlyWage = 5.64,
                ManHour = 174,
                UsualOvertime = 28,
                WeekendOvertime = 8,
                LegalOvertime = 0,
                BasePay = 981.36,
                OvertimePay = 252.88,
                FoodAllowance = 0,
                OtherAllowance = 0,
                GrossPay = 1234.24,
                BoardWages = 60,
                Utilities = 30,
                BadDebt = 90,
                RealWages = 1144.24,
                Remark = string.Empty
            });

            return list;
        }
    }
}
