# 快速入门

<a name="fb5456ad"></a>
### ExcelReport是什么？
ExcelReport是一个Excel模板渲染引擎。 它基于关注点分离的理念，将数据与表格样式、字体格式分离。<br />其中模板承载的表格样式、字体格式在可视化的情况下编辑。开发人员只需要绑定数据与目标标签的对应关系。ExcelReport就可以以数据驱动的方式渲染出目标报表。

<a name="b870eaea"></a>
### 模块组成
![image.png](https://cdn.nlark.com/yuque/0/2019/png/297115/1553490543172-faebb228-6700-4017-9e9b-27f9e7e0c9a1.png#align=left&display=inline&height=463&name=image.png&originHeight=463&originWidth=644&size=44745&status=done&width=644)<br />ExcelReport家族现在有四个成员。<br />ExcelReport负责报表的渲染逻辑。ExcelReport.Driver为ExcelReport提供了操作Excel文档的抽象接口。<br />ExcelReport.Driver.NPOI是使用NPOI对ExcelReport.Driver的实现。支持xls、xlsx两种格式的Excel文档。ExcelReport.Driver.CSV是针对csv格式的Excel文档对ExcelReport.Driver的实现。

<a name="bf553ad2"></a>
### 渲染模型
![image.png](https://cdn.nlark.com/yuque/0/2019/png/297115/1553578026046-7df71233-5118-422f-bdd6-85f026b181bd.png#align=left&display=inline&height=375&name=image.png&originHeight=375&originWidth=926&size=34989&status=done&width=926)<br />Template：模板承载的表格样式、字体格式、占位标签等。<br />Render：指定模板标签与数据的关系。<br />Data：注入模板的数据。整个渲染过程也是数据驱动渲染的。<br />Output：输出文件

<a name="3ddeeb49"></a>
### 入门示例
* 步骤一：新建入门项目QuickStart，并引入nuget包：

![image.png](https://cdn.nlark.com/yuque/0/2019/png/297115/1554000123061-20d4207c-58d3-4eb1-aeaa-0a8f7e3323ec.png#align=left&display=inline&height=69&name=image.png&originHeight=86&originWidth=936&size=4174&status=done&width=749)<br />![image.png](https://cdn.nlark.com/yuque/0/2019/png/297115/1554001512457-ed0bd004-a22d-47b2-ac0a-bdd673597144.png#align=left&display=inline&height=70&name=image.png&originHeight=87&originWidth=935&size=4451&status=done&width=748)

* 步骤二：创建并编辑模板

![image.png](https://cdn.nlark.com/yuque/0/2019/png/297115/1553996459416-086d8b4e-3c4c-4b4e-9774-53bd7f62c7ce.png#align=left&display=inline&height=698&name=image.png&originHeight=873&originWidth=449&size=35530&status=done&width=359)<br />![image.png](https://cdn.nlark.com/yuque/0/2019/png/297115/1553996053253-14c4f6f9-8d43-4866-af0b-d27b18e79af8.png#align=left&display=inline&height=330&name=image.png&originHeight=412&originWidth=1025&size=21989&status=done&width=820)

* 步骤三：编写代码

```csharp
internal class Program
{

       private static void Main(string[] args)
       {
           // 项目启动时，添加
           Configurator.Put(".xlsx", new WorkbookLoader());

           var num = 1;
           ExportHelper.ExportToLocal(@"templates\student.xlsx", "out.xlsx",
                   new SheetRenderer("Students",
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
           Console.WriteLine("finished!");
           Console.ReadKey();
       }
}
```

```csharp
public class StudentInfo
{
       public string Name { get; set; }
       public bool Gender { get; set; }
       public string Class { get; set; }
       public string RecordNo { get; set; }
       public string Phone { get; set; }
       public string Email { get; set; }
}
```

```csharp
public static class StudentLogic
{
       public static List<StudentInfo> GetList()
       {
           List<StudentInfo> list = new List<StudentInfo>();
           list.Add(new StudentInfo() { Class = "一班", Name = "XXX01", Gender = true, RecordNo = "YYY0001", Phone = "158******01", Email = "xxx01@live.cn" });
           list.Add(new StudentInfo() { Class = "二班", Name = "XXX02", Gender = false, RecordNo = "YYY0002", Phone = "158******02", Email = "xxx02@live.cn" });
           list.Add(new StudentInfo() { Class = "一班", Name = "XXX03", Gender = true, RecordNo = "YYY0003", Phone = "158******03", Email = "xxx03@live.cn" });
           list.Add(new StudentInfo() { Class = "一班", Name = "XXX04", Gender = true, RecordNo = "YYY0004", Phone = "158******04", Email = "xxx04@live.cn" });
           return list;
       }
}
```


* 输出结果
![image.png](https://cdn.nlark.com/yuque/0/2019/png/297115/1553998004321-0116c3dc-6980-4ef9-87b0-fd625c0e6602.png#align=left&display=inline&height=330&name=image.png&originHeight=412&originWidth=1025&size=24277&status=done&width=820)




### 更多文章
[ExcelReport文档](https://www.yuque.com/motse/excelreport)

### 联系我
![image.png](https://cdn.nlark.com/yuque/0/2019/png/297115/1556684663311-403a17ef-c19d-4f4e-8055-7f21044be0be.png)
