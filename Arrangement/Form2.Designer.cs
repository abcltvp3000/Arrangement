namespace Arrangement
{
    partial class Form2
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.arrangeTabCtrl = new System.Windows.Forms.TabControl();
            this.arrangeTab = new System.Windows.Forms.TabPage();
            this.arrangeExpBtn = new System.Windows.Forms.Button();
            this.arrangeGrid = new System.Windows.Forms.DataGridView();
            this.arrangeSchoolCode = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.arrangeSchool = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.arrangeName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.arrangeGroup = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.arrangeSchoolName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.arrangeJob = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.arrangeHead = new System.Windows.Forms.Label();
            this.arrangeTabCtrl.SuspendLayout();
            this.arrangeTab.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.arrangeGrid)).BeginInit();
            this.SuspendLayout();
            // 
            // arrangeTabCtrl
            // 
            this.arrangeTabCtrl.Controls.Add(this.arrangeTab);
            this.arrangeTabCtrl.Location = new System.Drawing.Point(12, 91);
            this.arrangeTabCtrl.Name = "arrangeTabCtrl";
            this.arrangeTabCtrl.SelectedIndex = 0;
            this.arrangeTabCtrl.Size = new System.Drawing.Size(936, 432);
            this.arrangeTabCtrl.TabIndex = 45;
            // 
            // arrangeTab
            // 
            this.arrangeTab.Controls.Add(this.arrangeExpBtn);
            this.arrangeTab.Controls.Add(this.arrangeGrid);
            this.arrangeTab.Location = new System.Drawing.Point(4, 24);
            this.arrangeTab.Name = "arrangeTab";
            this.arrangeTab.Padding = new System.Windows.Forms.Padding(3);
            this.arrangeTab.Size = new System.Drawing.Size(928, 404);
            this.arrangeTab.TabIndex = 0;
            this.arrangeTab.Text = "Theo trường";
            this.arrangeTab.UseVisualStyleBackColor = true;
            // 
            // arrangeExpBtn
            // 
            this.arrangeExpBtn.AutoSize = true;
            this.arrangeExpBtn.Location = new System.Drawing.Point(783, 3);
            this.arrangeExpBtn.Name = "arrangeExpBtn";
            this.arrangeExpBtn.Size = new System.Drawing.Size(145, 25);
            this.arrangeExpBtn.TabIndex = 21;
            this.arrangeExpBtn.Text = "Xuất kết quả ra file Excel";
            this.arrangeExpBtn.UseVisualStyleBackColor = true;
            this.arrangeExpBtn.Click += new System.EventHandler(this.button2_Click);
            // 
            // arrangeGrid
            // 
            this.arrangeGrid.AllowUserToAddRows = false;
            this.arrangeGrid.AllowUserToDeleteRows = false;
            this.arrangeGrid.AllowUserToOrderColumns = true;
            this.arrangeGrid.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.arrangeGrid.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.arrangeGrid.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.arrangeGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.arrangeGrid.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.arrangeSchoolCode,
            this.arrangeSchool,
            this.arrangeName,
            this.arrangeGroup,
            this.arrangeSchoolName,
            this.arrangeJob});
            this.arrangeGrid.Location = new System.Drawing.Point(0, 32);
            this.arrangeGrid.Name = "arrangeGrid";
            this.arrangeGrid.ReadOnly = true;
            this.arrangeGrid.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;
            this.arrangeGrid.RowTemplate.Height = 25;
            this.arrangeGrid.Size = new System.Drawing.Size(928, 372);
            this.arrangeGrid.TabIndex = 20;
            // 
            // arrangeSchoolCode
            // 
            this.arrangeSchoolCode.HeaderText = "Mã đơn vị";
            this.arrangeSchoolCode.Name = "arrangeSchoolCode";
            this.arrangeSchoolCode.ReadOnly = true;
            // 
            // arrangeSchool
            // 
            this.arrangeSchool.HeaderText = "Tên đơn vị";
            this.arrangeSchool.Name = "arrangeSchool";
            this.arrangeSchool.ReadOnly = true;
            // 
            // arrangeName
            // 
            this.arrangeName.HeaderText = "Họ và tên";
            this.arrangeName.Name = "arrangeName";
            this.arrangeName.ReadOnly = true;
            // 
            // arrangeGroup
            // 
            this.arrangeGroup.HeaderText = "Đoàn";
            this.arrangeGroup.Name = "arrangeGroup";
            this.arrangeGroup.ReadOnly = true;
            // 
            // arrangeSchoolName
            // 
            this.arrangeSchoolName.HeaderText = "Coi thi tại điểm thi";
            this.arrangeSchoolName.Name = "arrangeSchoolName";
            this.arrangeSchoolName.ReadOnly = true;
            // 
            // arrangeJob
            // 
            this.arrangeJob.HeaderText = "Chức vụ";
            this.arrangeJob.Name = "arrangeJob";
            this.arrangeJob.ReadOnly = true;
            // 
            // arrangeHead
            // 
            this.arrangeHead.AutoSize = true;
            this.arrangeHead.Font = new System.Drawing.Font("Segoe UI", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.arrangeHead.Location = new System.Drawing.Point(12, 30);
            this.arrangeHead.Name = "arrangeHead";
            this.arrangeHead.Size = new System.Drawing.Size(320, 30);
            this.arrangeHead.TabIndex = 43;
            this.arrangeHead.Text = "Sắp xếp phân công ngẫu nhiên";
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.ClientSize = new System.Drawing.Size(968, 564);
            this.Controls.Add(this.arrangeTabCtrl);
            this.Controls.Add(this.arrangeHead);
            this.Name = "Form2";
            this.Text = "Form2";
            this.arrangeTabCtrl.ResumeLayout(false);
            this.arrangeTab.ResumeLayout(false);
            this.arrangeTab.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.arrangeGrid)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private TabControl arrangeTabCtrl;
        private TabPage arrangeTab;
        private Button arrangeExpBtn;
        public DataGridView arrangeGrid;
        private Label arrangeHead;
        public DataGridViewTextBoxColumn arrangeSchoolCode;
        public DataGridViewTextBoxColumn arrangeSchool;
        public DataGridViewTextBoxColumn arrangeName;
        public DataGridViewTextBoxColumn arrangeGroup;
        public DataGridViewTextBoxColumn arrangeSchoolName;
        public DataGridViewTextBoxColumn arrangeJob;
    }
}