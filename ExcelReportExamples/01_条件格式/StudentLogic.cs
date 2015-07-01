using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace _01_条件格式
{
    public static class StudentLogic
    {
        public static List<StudentInfo> GetList()
        {
            List<StudentInfo> list = new List<StudentInfo>();

            list.Add(new StudentInfo() { Grade = 58, Name = "XXX01", Gender = true, RecordNo = "YYY0001", Phone = "158******01", Email = "xxx01@live.cn" });
            list.Add(new StudentInfo() { Grade = 73, Name = "XXX02", Gender = false, RecordNo = "YYY0002", Phone = "158******02", Email = "xxx02@live.cn" });
            list.Add(new StudentInfo() { Grade = 64, Name = "XXX03", Gender = true, RecordNo = "YYY0003", Phone = "158******03", Email = "xxx03@live.cn" });
            list.Add(new StudentInfo() { Grade = 99, Name = "XXX04", Gender = true, RecordNo = "YYY0004", Phone = "158******04", Email = "xxx04@live.cn" });
            list.Add(new StudentInfo() { Grade = 40, Name = "XXX05", Gender = true, RecordNo = "YYY0005", Phone = "158******05", Email = "xxx05@live.cn" });
            list.Add(new StudentInfo() { Grade = 96, Name = "XXX06", Gender = false, RecordNo = "YYY0006", Phone = "158******06", Email = "xxx06@live.cn" });
            list.Add(new StudentInfo() { Grade = 62, Name = "XXX07", Gender = true, RecordNo = "YYY0007", Phone = "158******07", Email = "xxx07@live.cn" });
            list.Add(new StudentInfo() { Grade = 70, Name = "XXX08", Gender = true, RecordNo = "YYY0008", Phone = "158******08", Email = "xxx08@live.cn" });

            return list;
        }
    }
}
