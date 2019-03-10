using ExcelReport;
using ExcelReport.Driver.CSV;
using ExcelReport.Renderers;
using System;

namespace _4.CSV示例
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            // 项目启动时，添加
            Configurator.Put(".csv", new WorkbookLoader());

            try
            {
                var num = 1;
                ExportHelper.ExportToLocal(@"Template\test.csv", "out.csv",
                        new SheetRenderer("test",
                            new RepeaterRenderer<StudentInfo>("Roster", StudentLogic.GetList(),
                                new ParameterRenderer<StudentInfo>("No", t => num++),
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