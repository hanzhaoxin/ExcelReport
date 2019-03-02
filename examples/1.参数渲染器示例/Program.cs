using ExcelReport;
using ExcelReport.Extends;
using ExcelReport.Renderers;
using System;
using System.Drawing;

namespace _1.参数渲染器示例
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            ExportHelper.ExportToLocal(@"Template\Template.xls", "out.xls",
                    new SheetRenderer("参数渲染示例",
                        new ParameterRenderer("String", "Hello World!"),
                        new ParameterRenderer("Boolean", true),
                        new ParameterRenderer("DateTime", DateTime.Now),
                        new ParameterRenderer("Double", 3.14),
                        new ParameterRenderer("Image", Image.FromFile("Image/C#高级编程.jpg"))
                        )
                    );
            Console.WriteLine("finished!");
        }
    }
}