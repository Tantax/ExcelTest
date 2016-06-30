namespace ExcelTest
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
            this.components = new System.ComponentModel.Container();
            this.textBoxExcelPath = new System.Windows.Forms.TextBox();
            this.buttonInExcel = new System.Windows.Forms.Button();
            this.buttonOutExcel = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.Id = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.paymentAccount = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.customerName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.adress = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.phoneNumber = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.gasType = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.factoryNumber = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.area = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.community = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.floor = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.userInformationBindingSource1 = new System.Windows.Forms.BindingSource(this.components);
            this.excelTestDBDataSet1 = new ExcelTest.ExcelTestDBDataSet1();
            this.userInformationBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.labelExcelUrl = new System.Windows.Forms.Label();
            this.userInformationTableAdapter1 = new ExcelTest.ExcelTestDBDataSet1TableAdapters.UserInformationTableAdapter();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.userInformationBindingSource1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.excelTestDBDataSet1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.userInformationBindingSource)).BeginInit();
            this.SuspendLayout();
            // 
            // textBoxExcelPath
            // 
            this.textBoxExcelPath.Location = new System.Drawing.Point(150, 14);
            this.textBoxExcelPath.Name = "textBoxExcelPath";
            this.textBoxExcelPath.Size = new System.Drawing.Size(249, 21);
            this.textBoxExcelPath.TabIndex = 0;
            this.textBoxExcelPath.DragDrop += new System.Windows.Forms.DragEventHandler(this.textBoxExcelPath_DragDrop);
            this.textBoxExcelPath.DragEnter += new System.Windows.Forms.DragEventHandler(this.textBoxExcelPath_DragEnter);
            // 
            // buttonInExcel
            // 
            this.buttonInExcel.Location = new System.Drawing.Point(406, 11);
            this.buttonInExcel.Name = "buttonInExcel";
            this.buttonInExcel.Size = new System.Drawing.Size(109, 23);
            this.buttonInExcel.TabIndex = 1;
            this.buttonInExcel.Text = "数据库导出Excel";
            this.buttonInExcel.UseVisualStyleBackColor = true;
            this.buttonInExcel.Click += new System.EventHandler(this.buttonInExcel_Click);
            // 
            // buttonOutExcel
            // 
            this.buttonOutExcel.Location = new System.Drawing.Point(521, 11);
            this.buttonOutExcel.Name = "buttonOutExcel";
            this.buttonOutExcel.Size = new System.Drawing.Size(103, 23);
            this.buttonOutExcel.TabIndex = 2;
            this.buttonOutExcel.Text = "Excel导入数据库";
            this.buttonOutExcel.UseVisualStyleBackColor = true;
            this.buttonOutExcel.Click += new System.EventHandler(this.buttonOutExcel_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.AutoGenerateColumns = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Id,
            this.paymentAccount,
            this.customerName,
            this.adress,
            this.phoneNumber,
            this.gasType,
            this.factoryNumber,
            this.area,
            this.community,
            this.floor});
            this.dataGridView1.DataSource = this.userInformationBindingSource1;
            this.dataGridView1.Location = new System.Drawing.Point(13, 41);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 23;
            this.dataGridView1.Size = new System.Drawing.Size(701, 327);
            this.dataGridView1.TabIndex = 3;
            // 
            // Id
            // 
            this.Id.DataPropertyName = "Id";
            this.Id.HeaderText = "Id";
            this.Id.Name = "Id";
            this.Id.ReadOnly = true;
            // 
            // paymentAccount
            // 
            this.paymentAccount.DataPropertyName = "PaymentAccount";
            this.paymentAccount.HeaderText = "缴费账户";
            this.paymentAccount.Name = "paymentAccount";
            // 
            // customerName
            // 
            this.customerName.DataPropertyName = "CustomerName";
            this.customerName.HeaderText = "用户姓名";
            this.customerName.Name = "customerName";
            // 
            // adress
            // 
            this.adress.DataPropertyName = "Adress";
            this.adress.HeaderText = "联系地址";
            this.adress.Name = "adress";
            // 
            // phoneNumber
            // 
            this.phoneNumber.DataPropertyName = "PhoneNumber";
            this.phoneNumber.HeaderText = "联系电话";
            this.phoneNumber.Name = "phoneNumber";
            // 
            // gasType
            // 
            this.gasType.DataPropertyName = "GasType";
            this.gasType.HeaderText = "用气类型";
            this.gasType.Name = "gasType";
            // 
            // factoryNumber
            // 
            this.factoryNumber.DataPropertyName = "FactoryNumber";
            this.factoryNumber.HeaderText = "出厂号码";
            this.factoryNumber.Name = "factoryNumber";
            // 
            // area
            // 
            this.area.DataPropertyName = "Area";
            this.area.HeaderText = "片区";
            this.area.Name = "area";
            // 
            // community
            // 
            this.community.DataPropertyName = "Community";
            this.community.HeaderText = "小区";
            this.community.Name = "community";
            // 
            // floor
            // 
            this.floor.DataPropertyName = "Floor";
            this.floor.HeaderText = "楼栋";
            this.floor.Name = "floor";
            // 
            // userInformationBindingSource1
            // 
            this.userInformationBindingSource1.DataMember = "UserInformation";
            this.userInformationBindingSource1.DataSource = this.excelTestDBDataSet1;
            // 
            // excelTestDBDataSet1
            // 
            this.excelTestDBDataSet1.DataSetName = "ExcelTestDBDataSet1";
            this.excelTestDBDataSet1.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // userInformationBindingSource
            // 
            this.userInformationBindingSource.DataMember = "UserInformation";
            this.userInformationBindingSource.DataSource = this.excelTestDBDataSet1;
            // 
            // labelExcelUrl
            // 
            this.labelExcelUrl.AutoSize = true;
            this.labelExcelUrl.Location = new System.Drawing.Point(13, 17);
            this.labelExcelUrl.Name = "labelExcelUrl";
            this.labelExcelUrl.Size = new System.Drawing.Size(131, 12);
            this.labelExcelUrl.TabIndex = 4;
            this.labelExcelUrl.Text = "Excel地址(可以拖拽)：";
            // 
            // userInformationTableAdapter1
            // 
            this.userInformationTableAdapter1.ClearBeforeFill = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(726, 380);
            this.Controls.Add(this.labelExcelUrl);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.buttonOutExcel);
            this.Controls.Add(this.buttonInExcel);
            this.Controls.Add(this.textBoxExcelPath);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.userInformationBindingSource1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.excelTestDBDataSet1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.userInformationBindingSource)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBoxExcelPath;
        private System.Windows.Forms.Button buttonInExcel;
        private System.Windows.Forms.Button buttonOutExcel;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.BindingSource userInformationBindingSource;
        private System.Windows.Forms.Label labelExcelUrl;
        private ExcelTestDBDataSet1 excelTestDBDataSet1;
        private System.Windows.Forms.BindingSource userInformationBindingSource1;
        private ExcelTestDBDataSet1TableAdapters.UserInformationTableAdapter userInformationTableAdapter1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Id;
        private System.Windows.Forms.DataGridViewTextBoxColumn paymentAccount;
        private System.Windows.Forms.DataGridViewTextBoxColumn customerName;
        private System.Windows.Forms.DataGridViewTextBoxColumn adress;
        private System.Windows.Forms.DataGridViewTextBoxColumn phoneNumber;
        private System.Windows.Forms.DataGridViewTextBoxColumn gasType;
        private System.Windows.Forms.DataGridViewTextBoxColumn factoryNumber;
        private System.Windows.Forms.DataGridViewTextBoxColumn area;
        private System.Windows.Forms.DataGridViewTextBoxColumn community;
        private System.Windows.Forms.DataGridViewTextBoxColumn floor;
    }
}

