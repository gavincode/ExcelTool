namespace ExcelTool
{
    partial class TableMapping
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(TableMapping));
            this.lblSheetColumnNames = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.btnADD = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.cmbDbTables = new System.Windows.Forms.ComboBox();
            this.rtxtMapping = new System.Windows.Forms.RichTextBox();
            this.lblSheetName = new System.Windows.Forms.Label();
            this.listViewExcelColumn = new System.Windows.Forms.ListView();
            this.listViewDbColumn = new System.Windows.Forms.ListView();
            this.txtSheeName = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnDelete = new System.Windows.Forms.Button();
            this.btnReset = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // lblSheetColumnNames
            // 
            this.lblSheetColumnNames.AutoSize = true;
            this.lblSheetColumnNames.Location = new System.Drawing.Point(512, 69);
            this.lblSheetColumnNames.Name = "lblSheetColumnNames";
            this.lblSheetColumnNames.Size = new System.Drawing.Size(95, 12);
            this.lblSheetColumnNames.TabIndex = 0;
            this.lblSheetColumnNames.Text = "Excel表单字段名";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(638, 129);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(65, 12);
            this.label1.TabIndex = 3;
            this.label1.Text = "--------->";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(743, 68);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(77, 12);
            this.label2.TabIndex = 0;
            this.label2.Text = "数据表字段名";
            // 
            // btnADD
            // 
            this.btnADD.Location = new System.Drawing.Point(640, 200);
            this.btnADD.Name = "btnADD";
            this.btnADD.Size = new System.Drawing.Size(63, 56);
            this.btnADD.TabIndex = 5;
            this.btnADD.Text = "添加";
            this.btnADD.UseVisualStyleBackColor = true;
            this.btnADD.Click += new System.EventHandler(this.btnADD_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(438, 26);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(47, 12);
            this.label3.TabIndex = 0;
            this.label3.Text = "映射表:";
            // 
            // cmbDbTables
            // 
            this.cmbDbTables.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbDbTables.FormattingEnabled = true;
            this.cmbDbTables.Location = new System.Drawing.Point(514, 21);
            this.cmbDbTables.Name = "cmbDbTables";
            this.cmbDbTables.Size = new System.Drawing.Size(316, 20);
            this.cmbDbTables.TabIndex = 2;
            this.cmbDbTables.SelectedIndexChanged += new System.EventHandler(this.cmbDbTables_SelectedIndexChanged);
            // 
            // rtxtMapping
            // 
            this.rtxtMapping.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.rtxtMapping.Location = new System.Drawing.Point(8, 89);
            this.rtxtMapping.Name = "rtxtMapping";
            this.rtxtMapping.ReadOnly = true;
            this.rtxtMapping.Size = new System.Drawing.Size(477, 496);
            this.rtxtMapping.TabIndex = 7;
            this.rtxtMapping.Text = "";
            // 
            // lblSheetName
            // 
            this.lblSheetName.AutoSize = true;
            this.lblSheetName.Location = new System.Drawing.Point(22, 27);
            this.lblSheetName.Name = "lblSheetName";
            this.lblSheetName.Size = new System.Drawing.Size(47, 12);
            this.lblSheetName.TabIndex = 8;
            this.lblSheetName.Text = "表单名:";
            // 
            // listViewExcelColumn
            // 
            this.listViewExcelColumn.BackColor = System.Drawing.Color.White;
            this.listViewExcelColumn.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.listViewExcelColumn.CheckBoxes = true;
            this.listViewExcelColumn.FullRowSelect = true;
            this.listViewExcelColumn.Location = new System.Drawing.Point(496, 89);
            this.listViewExcelColumn.Name = "listViewExcelColumn";
            this.listViewExcelColumn.Size = new System.Drawing.Size(136, 496);
            this.listViewExcelColumn.TabIndex = 10;
            this.listViewExcelColumn.UseCompatibleStateImageBehavior = false;
            this.listViewExcelColumn.View = System.Windows.Forms.View.List;
            this.listViewExcelColumn.ItemChecked += new System.Windows.Forms.ItemCheckedEventHandler(this.listViewExcelColumn_ItemChecked);
            // 
            // listViewDbColumn
            // 
            this.listViewDbColumn.AllowDrop = true;
            this.listViewDbColumn.BackColor = System.Drawing.Color.White;
            this.listViewDbColumn.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.listViewDbColumn.CheckBoxes = true;
            this.listViewDbColumn.Location = new System.Drawing.Point(714, 89);
            this.listViewDbColumn.Name = "listViewDbColumn";
            this.listViewDbColumn.Size = new System.Drawing.Size(136, 496);
            this.listViewDbColumn.TabIndex = 11;
            this.listViewDbColumn.UseCompatibleStateImageBehavior = false;
            this.listViewDbColumn.View = System.Windows.Forms.View.List;
            this.listViewDbColumn.ItemChecked += new System.Windows.Forms.ItemCheckedEventHandler(this.listViewDbColumn_ItemChecked);
            // 
            // txtSheeName
            // 
            this.txtSheeName.Enabled = false;
            this.txtSheeName.Location = new System.Drawing.Point(75, 23);
            this.txtSheeName.Name = "txtSheeName";
            this.txtSheeName.Size = new System.Drawing.Size(237, 21);
            this.txtSheeName.TabIndex = 12;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(341, 27);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(65, 12);
            this.label4.TabIndex = 13;
            this.label4.Text = "--------->";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(12, 68);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(59, 12);
            this.label5.TabIndex = 14;
            this.label5.Text = "映射关系:";
            // 
            // btnSave
            // 
            this.btnSave.Enabled = false;
            this.btnSave.Location = new System.Drawing.Point(640, 285);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(63, 56);
            this.btnSave.TabIndex = 15;
            this.btnSave.Text = "保存";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnDelete
            // 
            this.btnDelete.Location = new System.Drawing.Point(640, 373);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(63, 56);
            this.btnDelete.TabIndex = 16;
            this.btnDelete.Text = "删除";
            this.btnDelete.UseVisualStyleBackColor = true;
            this.btnDelete.Visible = false;
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            // 
            // btnReset
            // 
            this.btnReset.Enabled = false;
            this.btnReset.Location = new System.Drawing.Point(640, 372);
            this.btnReset.Name = "btnReset";
            this.btnReset.Size = new System.Drawing.Size(63, 56);
            this.btnReset.TabIndex = 17;
            this.btnReset.Text = "重置";
            this.btnReset.UseVisualStyleBackColor = true;
            this.btnReset.Visible = false;
            this.btnReset.Click += new System.EventHandler(this.btnReset_Click);
            // 
            // TableMapping
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(861, 589);
            this.Controls.Add(this.btnReset);
            this.Controls.Add(this.btnDelete);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.txtSheeName);
            this.Controls.Add(this.listViewDbColumn);
            this.Controls.Add(this.listViewExcelColumn);
            this.Controls.Add(this.lblSheetName);
            this.Controls.Add(this.rtxtMapping);
            this.Controls.Add(this.btnADD);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cmbDbTables);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.lblSheetColumnNames);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "TableMapping";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "字段关系映射";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblSheetColumnNames;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnADD;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cmbDbTables;
        private System.Windows.Forms.RichTextBox rtxtMapping;
        private System.Windows.Forms.Label lblSheetName;
        private System.Windows.Forms.ListView listViewExcelColumn;
        private System.Windows.Forms.ListView listViewDbColumn;
        private System.Windows.Forms.TextBox txtSheeName;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnDelete;
        private System.Windows.Forms.Button btnReset;
    }
}