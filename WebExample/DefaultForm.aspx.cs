using System;
using System.Collections.Generic;
using ExcelReport;

namespace WebExample
{
    public partial class DefaultForm : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
        }

        protected void btnExport_Click(object sender, EventArgs e)
        {
            ExcelReport.ParameterCollection collection = new ExcelReport.ParameterCollection();
            collection.Load(Server.MapPath(@"Template\Template.xml"));

            List<ElementFormatter> formatters = new List<ElementFormatter>();
            formatters.Add(new CellFormatter(collection["GradeDetail", "Dept"], "某某学院"));
            formatters.Add(new CellFormatter(collection["GradeDetail", "Class"], "某某班级"));
            formatters.Add(new CellFormatter(collection["GradeDetail", "StudNo"], "2009*****"));
            formatters.Add(new CellFormatter(collection["GradeDetail", "StudName"], "韩兆新"));
            formatters.Add(new CellFormatter(collection["GradeDetail", "ExportDate"], DateTime.Now));

            List<GradeInfo> gradeInfoList = new List<GradeInfo>();
            gradeInfoList.Add(new GradeInfo() { CGPA = 18, CourseID = "KC-0001", CourseName = "高等数学", CourseType = "理论课", Credit = 6, EvaluationMode = "考试", GainCredit = 6, GPA = 3, Grade = 86, StudyNature = "初修", Type = "必修课" });
            gradeInfoList.Add(new GradeInfo() { CGPA = 2, CourseID = "KC-0002", CourseName = "计算机应用基础", CourseType = "理论课", Credit = 2, EvaluationMode = "考试", GainCredit = 2, GPA = 1, Grade = 93, StudyNature = "初修", Type = "必修课" });
            gradeInfoList.Add(new GradeInfo() { CGPA = 9, CourseID = "KC-0003", CourseName = "C程序设计", CourseType = "理论课", Credit = 3, EvaluationMode = "考试", GainCredit = 3, GPA = 3, Grade = 97, StudyNature = "初修", Type = "必修课", Remark = "备注信息" });

            formatters.Add(new TableFormatter<GradeInfo>(collection["GradeDetail", "CourseID"].X, gradeInfoList,
                new TableColumnInfo<GradeInfo>(collection["GradeDetail", "CGPA"].Y, t => t.CGPA),
                new TableColumnInfo<GradeInfo>(collection["GradeDetail", "CourseID"].Y, t => t.CourseID),
                new TableColumnInfo<GradeInfo>(collection["GradeDetail", "CourseName"].Y, t => t.CourseName),
                new TableColumnInfo<GradeInfo>(collection["GradeDetail", "CourseType"].Y, t => t.CourseType),
                new TableColumnInfo<GradeInfo>(collection["GradeDetail", "Credit"].Y, t => t.Credit),
                new TableColumnInfo<GradeInfo>(collection["GradeDetail", "EvaluationMode"].Y, t => t.EvaluationMode),
                new TableColumnInfo<GradeInfo>(collection["GradeDetail", "GainCredit"].Y, t => t.GainCredit),
                new TableColumnInfo<GradeInfo>(collection["GradeDetail", "GPA"].Y, t => t.GPA),
                new TableColumnInfo<GradeInfo>(collection["GradeDetail", "Grade"].Y, t => t.Grade),
                new TableColumnInfo<GradeInfo>(collection["GradeDetail", "Remark"].Y, t => t.Remark),
                new TableColumnInfo<GradeInfo>(collection["GradeDetail", "StudyNature"].Y, t => t.StudyNature),
                new TableColumnInfo<GradeInfo>(collection["GradeDetail", "Type"].Y, t => t.Type)
                ));
            Export.ExportToWeb(Server.MapPath(@"Template\Template.xls"), "GradeDetail",
                new SheetFormatterContainer("GradeDetail", formatters));
        }
    }
}