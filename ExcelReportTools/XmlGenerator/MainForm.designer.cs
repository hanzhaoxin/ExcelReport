namespace XmlGenerator
{
    partial class MainForm
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.lbExcelTemplatePath = new System.Windows.Forms.Label();
            this.txtExcelTemplatePath = new System.Windows.Forms.TextBox();
            this.btnSelecttxtExcelTemplate = new System.Windows.Forms.Button();
            this.btnFillTemplate = new System.Windows.Forms.Button();
            this.lbStep_1 = new System.Windows.Forms.Label();
            this.lbStep_2 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // lbExcelTemplatePath
            // 
            this.lbExcelTemplatePath.AutoSize = true;
            this.lbExcelTemplatePath.Location = new System.Drawing.Point(13, 49);
            this.lbExcelTemplatePath.Name = "lbExcelTemplatePath";
            this.lbExcelTemplatePath.Size = new System.Drawing.Size(95, 12);
            this.lbExcelTemplatePath.TabIndex = 0;
            this.lbExcelTemplatePath.Text = "Excel模板文件：";
            // 
            // txtExcelTemplatePath
            // 
            this.txtExcelTemplatePath.AllowDrop = true;
            this.txtExcelTemplatePath.Location = new System.Drawing.Point(115, 45);
            this.txtExcelTemplatePath.Name = "txtExcelTemplatePath";
            this.txtExcelTemplatePath.Size = new System.Drawing.Size(347, 21);
            this.txtExcelTemplatePath.TabIndex = 1;
            this.txtExcelTemplatePath.DragDrop += new System.Windows.Forms.DragEventHandler(this.txtExcelTemplatePath_DragDrop);
            this.txtExcelTemplatePath.DragEnter += new System.Windows.Forms.DragEventHandler(this.txtExcelTemplatePath_DragEnter);
            // 
            // btnSelecttxtExcelTemplate
            // 
            this.btnSelecttxtExcelTemplate.Location = new System.Drawing.Point(468, 44);
            this.btnSelecttxtExcelTemplate.Name = "btnSelecttxtExcelTemplate";
            this.btnSelecttxtExcelTemplate.Size = new System.Drawing.Size(75, 23);
            this.btnSelecttxtExcelTemplate.TabIndex = 2;
            this.btnSelecttxtExcelTemplate.Text = "选择...";
            this.btnSelecttxtExcelTemplate.UseVisualStyleBackColor = true;
            this.btnSelecttxtExcelTemplate.Click += new System.EventHandler(this.btnSelecttxtExcelTemplate_Click);
            // 
            // btnFillTemplate
            // 
            this.btnFillTemplate.Location = new System.Drawing.Point(12, 126);
            this.btnFillTemplate.Name = "btnFillTemplate";
            this.btnFillTemplate.Size = new System.Drawing.Size(530, 23);
            this.btnFillTemplate.TabIndex = 3;
            this.btnFillTemplate.Text = "生成填充模板规则文件（.XML）";
            this.btnFillTemplate.UseVisualStyleBackColor = true;
            this.btnFillTemplate.Click += new System.EventHandler(this.btnFillTemplate_Click);
            // 
            // lbStep_1
            // 
            this.lbStep_1.AutoSize = true;
            this.lbStep_1.ForeColor = System.Drawing.SystemColors.MenuHighlight;
            this.lbStep_1.Location = new System.Drawing.Point(13, 13);
            this.lbStep_1.Name = "lbStep_1";
            this.lbStep_1.Size = new System.Drawing.Size(491, 12);
            this.lbStep_1.TabIndex = 4;
            this.lbStep_1.Text = "第一步：点击“选择...”按钮，选择Excel模板文件；或将Excel模板文件拖入下面文本框。";
            // 
            // lbStep_2
            // 
            this.lbStep_2.AutoSize = true;
            this.lbStep_2.ForeColor = System.Drawing.SystemColors.MenuHighlight;
            this.lbStep_2.Location = new System.Drawing.Point(13, 101);
            this.lbStep_2.Name = "lbStep_2";
            this.lbStep_2.Size = new System.Drawing.Size(353, 12);
            this.lbStep_2.TabIndex = 5;
            this.lbStep_2.Text = "第二步：点击“生成模板规则文件(.XML)”按钮，生成规则文件。";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(555, 161);
            this.Controls.Add(this.lbStep_2);
            this.Controls.Add(this.lbStep_1);
            this.Controls.Add(this.btnFillTemplate);
            this.Controls.Add(this.btnSelecttxtExcelTemplate);
            this.Controls.Add(this.txtExcelTemplatePath);
            this.Controls.Add(this.lbExcelTemplatePath);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "MainForm";
            this.Text = "模板填充规则文件生成工具";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lbExcelTemplatePath;
        private System.Windows.Forms.TextBox txtExcelTemplatePath;
        private System.Windows.Forms.Button btnSelecttxtExcelTemplate;
        private System.Windows.Forms.Button btnFillTemplate;
        private System.Windows.Forms.Label lbStep_1;
        private System.Windows.Forms.Label lbStep_2;
    }
}

