namespace 薪资发放邮件系统
{
    partial class Form1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.sendPathText = new System.Windows.Forms.TextBox();
            this.selectSendBtn = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.rePathText = new System.Windows.Forms.TextBox();
            this.selectReceBtn = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.salaryText = new System.Windows.Forms.TextBox();
            this.selectSendExcelBtn = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.excelType = new System.Windows.Forms.ComboBox();
            this.label6 = new System.Windows.Forms.Label();
            this.eBody = new System.Windows.Forms.TextBox();
            this.sendBtn = new System.Windows.Forms.Button();
            this.cancel = new System.Windows.Forms.Button();
            this.label7 = new System.Windows.Forms.Label();
            this.infoOne = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.infoTwo = new System.Windows.Forms.Label();
            this.infoThree = new System.Windows.Forms.Label();
            this.infoFour = new System.Windows.Forms.Label();
            this.infoFive = new System.Windows.Forms.Label();
            this.skinEngine1 = new Sunisoft.IrisSkin.SkinEngine(((System.ComponentModel.Component)(this)));
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(309, 48);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(149, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "作者：汪浩 肖建茂 方智峰";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(19, 79);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(89, 12);
            this.label2.TabIndex = 1;
            this.label2.Text = "发送邮箱文件：";
            // 
            // sendPathText
            // 
            this.sendPathText.Location = new System.Drawing.Point(114, 78);
            this.sendPathText.Name = "sendPathText";
            this.sendPathText.Size = new System.Drawing.Size(167, 21);
            this.sendPathText.TabIndex = 2;
            // 
            // selectSendBtn
            // 
            this.selectSendBtn.Location = new System.Drawing.Point(310, 74);
            this.selectSendBtn.Name = "selectSendBtn";
            this.selectSendBtn.Size = new System.Drawing.Size(130, 23);
            this.selectSendBtn.TabIndex = 3;
            this.selectSendBtn.Text = "请选择发送邮箱";
            this.selectSendBtn.UseVisualStyleBackColor = true;
            this.selectSendBtn.Click += new System.EventHandler(this.selectSendBtn_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(19, 117);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(89, 12);
            this.label3.TabIndex = 4;
            this.label3.Text = "接收邮箱文件：";
            // 
            // rePathText
            // 
            this.rePathText.Location = new System.Drawing.Point(114, 114);
            this.rePathText.Name = "rePathText";
            this.rePathText.Size = new System.Drawing.Size(167, 21);
            this.rePathText.TabIndex = 5;
            // 
            // selectReceBtn
            // 
            this.selectReceBtn.Location = new System.Drawing.Point(310, 112);
            this.selectReceBtn.Name = "selectReceBtn";
            this.selectReceBtn.Size = new System.Drawing.Size(130, 23);
            this.selectReceBtn.TabIndex = 6;
            this.selectReceBtn.Text = "请选择接收邮箱";
            this.selectReceBtn.UseVisualStyleBackColor = true;
            this.selectReceBtn.Click += new System.EventHandler(this.selectReceBtn_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(43, 199);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(65, 12);
            this.label4.TabIndex = 7;
            this.label4.Text = "发送文件：";
            // 
            // salaryText
            // 
            this.salaryText.Location = new System.Drawing.Point(114, 196);
            this.salaryText.Name = "salaryText";
            this.salaryText.Size = new System.Drawing.Size(168, 21);
            this.salaryText.TabIndex = 8;
            // 
            // selectSendExcelBtn
            // 
            this.selectSendExcelBtn.Location = new System.Drawing.Point(311, 194);
            this.selectSendExcelBtn.Name = "selectSendExcelBtn";
            this.selectSendExcelBtn.Size = new System.Drawing.Size(130, 23);
            this.selectSendExcelBtn.TabIndex = 9;
            this.selectSendExcelBtn.Text = "请选择发送文件";
            this.selectSendExcelBtn.UseVisualStyleBackColor = true;
            this.selectSendExcelBtn.Click += new System.EventHandler(this.selectSendExcelBtn_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(19, 159);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(89, 12);
            this.label5.TabIndex = 10;
            this.label5.Text = "发送文件类型：";
            // 
            // excelType
            // 
            this.excelType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.excelType.FormattingEnabled = true;
            this.excelType.Items.AddRange(new object[] {
            "工资明细",
            "无限制(第一行为人员代码)"});
            this.excelType.Location = new System.Drawing.Point(114, 154);
            this.excelType.Name = "excelType";
            this.excelType.Size = new System.Drawing.Size(168, 20);
            this.excelType.TabIndex = 12;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(42, 246);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(65, 12);
            this.label6.TabIndex = 13;
            this.label6.Text = "邮件内容：";
            // 
            // eBody
            // 
            this.eBody.Location = new System.Drawing.Point(114, 244);
            this.eBody.Multiline = true;
            this.eBody.Name = "eBody";
            this.eBody.Size = new System.Drawing.Size(327, 88);
            this.eBody.TabIndex = 14;
            // 
            // sendBtn
            // 
            this.sendBtn.Location = new System.Drawing.Point(112, 441);
            this.sendBtn.Name = "sendBtn";
            this.sendBtn.Size = new System.Drawing.Size(75, 23);
            this.sendBtn.TabIndex = 15;
            this.sendBtn.Text = "发送";
            this.sendBtn.UseVisualStyleBackColor = true;
            this.sendBtn.Click += new System.EventHandler(this.sendBtn_Click);
            // 
            // cancel
            // 
            this.cancel.Location = new System.Drawing.Point(310, 441);
            this.cancel.Name = "cancel";
            this.cancel.Size = new System.Drawing.Size(75, 23);
            this.cancel.TabIndex = 16;
            this.cancel.Text = "取消";
            this.cancel.UseVisualStyleBackColor = true;
            this.cancel.Click += new System.EventHandler(this.cancel_Click);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(19, 348);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(89, 12);
            this.label7.TabIndex = 17;
            this.label7.Text = "发送总体结果：";
            // 
            // infoOne
            // 
            this.infoOne.AutoSize = true;
            this.infoOne.Location = new System.Drawing.Point(111, 348);
            this.infoOne.Name = "infoOne";
            this.infoOne.Size = new System.Drawing.Size(233, 12);
            this.infoOne.TabIndex = 19;
            this.infoOne.Text = "总共发送了0条信息，已发送成功了0条信息";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("宋体", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label9.Location = new System.Drawing.Point(55, 14);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(366, 20);
            this.label9.TabIndex = 20;
            this.label9.Text = "江西师范大学人事处工资邮件发放系统";
            // 
            // infoTwo
            // 
            this.infoTwo.AutoSize = true;
            this.infoTwo.Location = new System.Drawing.Point(112, 365);
            this.infoTwo.Name = "infoTwo";
            this.infoTwo.Size = new System.Drawing.Size(233, 12);
            this.infoTwo.TabIndex = 22;
            this.infoTwo.Text = "总共发送了0条信息，已发送成功了0条信息";
            // 
            // infoThree
            // 
            this.infoThree.AutoSize = true;
            this.infoThree.Location = new System.Drawing.Point(111, 382);
            this.infoThree.Name = "infoThree";
            this.infoThree.Size = new System.Drawing.Size(233, 12);
            this.infoThree.TabIndex = 23;
            this.infoThree.Text = "总共发送了0条信息，已发送成功了0条信息";
            // 
            // infoFour
            // 
            this.infoFour.AutoSize = true;
            this.infoFour.Location = new System.Drawing.Point(111, 396);
            this.infoFour.Name = "infoFour";
            this.infoFour.Size = new System.Drawing.Size(233, 12);
            this.infoFour.TabIndex = 24;
            this.infoFour.Text = "总共发送了0条信息，已发送成功了0条信息";
            // 
            // infoFive
            // 
            this.infoFive.AutoSize = true;
            this.infoFive.Location = new System.Drawing.Point(112, 413);
            this.infoFive.Name = "infoFive";
            this.infoFive.Size = new System.Drawing.Size(233, 12);
            this.infoFive.TabIndex = 25;
            this.infoFive.Text = "总共发送了0条信息，已发送成功了0条信息";
            // 
            // skinEngine1
            // 
            this.skinEngine1.SerialNumber = "";
            this.skinEngine1.SkinFile = null;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(467, 469);
            this.Controls.Add(this.infoFive);
            this.Controls.Add(this.infoFour);
            this.Controls.Add(this.infoThree);
            this.Controls.Add(this.infoTwo);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.infoOne);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.cancel);
            this.Controls.Add(this.sendBtn);
            this.Controls.Add(this.eBody);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.excelType);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.selectSendExcelBtn);
            this.Controls.Add(this.salaryText);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.selectReceBtn);
            this.Controls.Add(this.rePathText);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.selectSendBtn);
            this.Controls.Add(this.sendPathText);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "江西师范大学人事处工资邮件发放系统";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing_1);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button selectSendBtn;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox rePathText;
        private System.Windows.Forms.Button selectReceBtn;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox salaryText;
        private System.Windows.Forms.Button selectSendExcelBtn;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ComboBox excelType;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox eBody;
        private System.Windows.Forms.Button sendBtn;
        private System.Windows.Forms.Button cancel;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label infoOne;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox sendPathText;
        private System.Windows.Forms.Label infoTwo;
        private System.Windows.Forms.Label infoThree;
        private System.Windows.Forms.Label infoFour;
        private System.Windows.Forms.Label infoFive;
        private Sunisoft.IrisSkin.SkinEngine skinEngine1;
    }
}

