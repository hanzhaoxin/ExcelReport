using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ExcelReport;

namespace RepeaterFormatterExample
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDlg = new SaveFileDialog();
            saveFileDlg.Filter = "Excel 2003文件|*.xls|Excel 2007文件|*.xlsx";

            if (DialogResult.OK.Equals(saveFileDlg.ShowDialog()))
            {
                //实例化一个参数容器，并加载模板填充规则文件
                ParameterCollection collection = new ParameterCollection();
                collection.Load(@"Template\Template.xml");

                //实例化一个元素格式化器列表
                List<ElementFormatter> formatters = new List<ElementFormatter>();

                List<PostInfo> postInfoList = new List<PostInfo>();
                postInfoList.Add(new PostInfo()
                {
                    Client = "客户一",
                    From = "来件地址一",
                    FromTel = "来件人电话一",
                    ReturnUnit = "退件单位一",
                    To = "收件地址一",
                    ToTel = "收件人电话一"
                });
                postInfoList.Add(new PostInfo()
                {
                    Client = "客户二",
                    From = "来件地址二",
                    FromTel = "来件人电话二",
                    ReturnUnit = "退件单位二",
                    To = "收件地址二",
                    ToTel = "收件人电话二"
                });
                postInfoList.Add(new PostInfo()
                {
                    Client = "客户三",
                    From = "来件地址三",
                    FromTel = "来件人电话三",
                    ReturnUnit = "退件单位三",
                    To = "收件地址三",
                    ToTel = "收件人电话三"
                });
                

                //添加一个Table格式化器
                formatters.Add(new RepeaterFormatter<PostInfo>(collection["Sheet1", "rptPostInfoList_Start"], collection["Sheet1", "rptPostInfoList_End"],
                    postInfoList,
                    new RepeaterCellInfo<PostInfo>(collection["Sheet1", "Client"], t => t.Client),
                    new RepeaterCellInfo<PostInfo>(collection["Sheet1", "From"], t => t.From),
                    new RepeaterCellInfo<PostInfo>(collection["Sheet1", "FromTel"], t => t.FromTel),
                    new RepeaterCellInfo<PostInfo>(collection["Sheet1", "To"], t => t.To),
                    new RepeaterCellInfo<PostInfo>(collection["Sheet1", "ToTel"], t => t.ToTel),
                    new RepeaterCellInfo<PostInfo>(collection["Sheet1", "ReturnUnit"], t => t.ReturnUnit)
                    ));
                //导出文件到本地
                Export.ExportToLocal(@"Template\Template.xls", saveFileDlg.FileName,
                    new SheetFormatterContainer("Sheet1", formatters));
            }
        }
    }
}
