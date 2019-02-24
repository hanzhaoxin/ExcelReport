using ExcelReport;
using ExcelReport.Renderers;
using System;

namespace _2.单行重复渲染示例
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            var num = 1;
            ExportHelper.ExportToLocal(@"Template\Template.xls", "out.xls",
                    new SheetRenderer("学生名册",
                        new RepeaterRenderer<StudentInfo>("Roster", StudentLogic.GetList(),
                            new ParameterRenderer<StudentInfo>("No", t => num++),
                            new ParameterRenderer<StudentInfo>("Name", t => t.Name),
                            new ParameterRenderer<StudentInfo>("Gender", t => t.Gender ? "男" : "女"),
                            new ParameterRenderer<StudentInfo>("Class", t => t.Class),
                            new ParameterRenderer<StudentInfo>("RecordNo", t => t.RecordNo),
                            new ParameterRenderer<StudentInfo>("Phone", t => t.Phone),
                            new ParameterRenderer<StudentInfo>("Email", t => t.Email)
                            )
                        )
                    );
            Console.WriteLine("finished!");
        }
    }
}