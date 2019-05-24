using ExcelReport;
using ExcelReport.Driver.NPOI;
using ExcelReport.Renderers;
using System;

namespace _2.单行重复渲染示例
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            // 项目启动时，添加
            Configurator.Put(".xls", new WorkbookLoader());

            try
            {
                var num = 1;
                ExportHelper.ExportToLocal(@"Template\Template.xls", "out.xls",
                        new SheetRenderer("学生名册",
                            new RepeaterRenderer<StudentInfo>("Roster", StudentLogic.GetList(),
                                new ParameterRenderer<StudentInfo>("No1", t => num++),
                                new ParameterRenderer<StudentInfo>("Name", t => t.Name),
                                new ParameterRenderer<StudentInfo>("Gender", t => t.Gender ? "男" : "女"),
                                new ParameterRenderer<StudentInfo>("Class", t => t.Class),
                                new ParameterRenderer<StudentInfo>("RecordNo", t => t.RecordNo),
                                new ParameterRenderer<StudentInfo>("Phone", t => t.Phone),
                                new ParameterRenderer<StudentInfo>("Email", t => t.Email)
                                ),
                             new ParameterRenderer("Author", "hzx")
                            )
                        );
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            Console.WriteLine("finished!");
            Console.ReadKey();
        }
    }
}