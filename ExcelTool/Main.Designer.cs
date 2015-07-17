namespace ExcelTool
{
    partial class Main
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Main));
            this.btnExpertExcel = new System.Windows.Forms.Button();
            this.tabControlSheetInfo = new System.Windows.Forms.TabControl();
            this.label1 = new System.Windows.Forms.Label();
            this.txtSelectExcel = new System.Windows.Forms.TextBox();
            this.btnImportBatchSheets = new System.Windows.Forms.Button();
            this.pgbBatchImport = new System.Windows.Forms.ProgressBar();
            this.lblImportTableName = new System.Windows.Forms.Label();
            this.lblSheetInfo = new System.Windows.Forms.Label();
            this.tabControlMain = new System.Windows.Forms.TabControl();
            this.tabPageImport = new System.Windows.Forms.TabPage();
            this.lblLoadOverTimeTip = new System.Windows.Forms.Label();
            this.ckbSelectAllNodes = new System.Windows.Forms.CheckBox();
            this.txtFind = new System.Windows.Forms.TextBox();
            this.treeViewExcels = new System.Windows.Forms.TreeView();
            this.tabPageExport = new System.Windows.Forms.TabPage();
            this.txtFindExportTable = new System.Windows.Forms.TextBox();
            this.ListviewTableNames = new System.Windows.Forms.ListView();
            this.tabControlDbData = new System.Windows.Forms.TabControl();
            this.tabPageLog = new System.Windows.Forms.TabPage();
            this.rtxtLog = new System.Windows.Forms.RichTextBox();
            this.cmsDeleteLog = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.toolStripMenuItemClear = new System.Windows.Forms.ToolStripMenuItem();
            this.tabPageSQL = new System.Windows.Forms.TabPage();
            this.rtxtSQL = new System.Windows.Forms.RichTextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtConnetingString = new System.Windows.Forms.TextBox();
            this.lblDbAccess = new System.Windows.Forms.Label();
            this.btnTestDbConnection = new System.Windows.Forms.Button();
            this.cmsSheetNode = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.toolStripMenuItemCopy = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.toolStripMenuMapping = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.toolStripTextBoxComment = new System.Windows.Forms.ToolStripTextBox();
            this.cmsOpenRoot = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.toolStripMenuItemOpenRoot = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSyncConfig = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripUpdateTime = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripReboot = new System.Windows.Forms.ToolStripMenuItem();
            this.tabControlMain.SuspendLayout();
            this.tabPageImport.SuspendLayout();
            this.tabPageExport.SuspendLayout();
            this.tabPageLog.SuspendLayout();
            this.cmsDeleteLog.SuspendLayout();
            this.tabPageSQL.SuspendLayout();
            this.cmsSheetNode.SuspendLayout();
            this.cmsOpenRoot.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnExpertExcel
            // 
            resources.ApplyResources(this.btnExpertExcel, "btnExpertExcel");
            this.btnExpertExcel.Name = "btnExpertExcel";
            this.btnExpertExcel.UseVisualStyleBackColor = true;
            this.btnExpertExcel.Click += new System.EventHandler(this.btnExportExcel_Click);
            // 
            // tabControlSheetInfo
            // 
            this.tabControlSheetInfo.AllowDrop = true;
            resources.ApplyResources(this.tabControlSheetInfo, "tabControlSheetInfo");
            this.tabControlSheetInfo.Name = "tabControlSheetInfo";
            this.tabControlSheetInfo.SelectedIndex = 0;
            this.tabControlSheetInfo.SelectedIndexChanged += new System.EventHandler(this.tabControlSheetInfo_SelectedIndexChanged);
            // 
            // label1
            // 
            resources.ApplyResources(this.label1, "label1");
            this.label1.Name = "label1";
            // 
            // txtSelectExcel
            // 
            this.txtSelectExcel.ForeColor = System.Drawing.Color.Gray;
            resources.ApplyResources(this.txtSelectExcel, "txtSelectExcel");
            this.txtSelectExcel.Name = "txtSelectExcel";
            this.txtSelectExcel.MouseClick += new System.Windows.Forms.MouseEventHandler(this.txtSelectExcel_MouseClick);
            // 
            // btnImportBatchSheets
            // 
            resources.ApplyResources(this.btnImportBatchSheets, "btnImportBatchSheets");
            this.btnImportBatchSheets.Name = "btnImportBatchSheets";
            this.btnImportBatchSheets.UseVisualStyleBackColor = true;
            this.btnImportBatchSheets.Click += new System.EventHandler(this.btnImportBatchSheets_Click);
            // 
            // pgbBatchImport
            // 
            resources.ApplyResources(this.pgbBatchImport, "pgbBatchImport");
            this.pgbBatchImport.Name = "pgbBatchImport";
            // 
            // lblImportTableName
            // 
            resources.ApplyResources(this.lblImportTableName, "lblImportTableName");
            this.lblImportTableName.Name = "lblImportTableName";
            // 
            // lblSheetInfo
            // 
            resources.ApplyResources(this.lblSheetInfo, "lblSheetInfo");
            this.lblSheetInfo.ForeColor = System.Drawing.Color.Red;
            this.lblSheetInfo.Name = "lblSheetInfo";
            // 
            // tabControlMain
            // 
            this.tabControlMain.Controls.Add(this.tabPageImport);
            this.tabControlMain.Controls.Add(this.tabPageExport);
            this.tabControlMain.Controls.Add(this.tabPageLog);
            this.tabControlMain.Controls.Add(this.tabPageSQL);
            resources.ApplyResources(this.tabControlMain, "tabControlMain");
            this.tabControlMain.Name = "tabControlMain";
            this.tabControlMain.SelectedIndex = 0;
            this.tabControlMain.SelectedIndexChanged += new System.EventHandler(this.tabControlMain_SelectedIndexChanged);
            // 
            // tabPageImport
            // 
            this.tabPageImport.Controls.Add(this.lblLoadOverTimeTip);
            this.tabPageImport.Controls.Add(this.ckbSelectAllNodes);
            this.tabPageImport.Controls.Add(this.txtFind);
            this.tabPageImport.Controls.Add(this.treeViewExcels);
            this.tabPageImport.Controls.Add(this.label1);
            this.tabPageImport.Controls.Add(this.txtSelectExcel);
            this.tabPageImport.Controls.Add(this.btnImportBatchSheets);
            this.tabPageImport.Controls.Add(this.tabControlSheetInfo);
            this.tabPageImport.Controls.Add(this.lblSheetInfo);
            this.tabPageImport.Controls.Add(this.pgbBatchImport);
            this.tabPageImport.Controls.Add(this.lblImportTableName);
            resources.ApplyResources(this.tabPageImport, "tabPageImport");
            this.tabPageImport.Name = "tabPageImport";
            this.tabPageImport.UseVisualStyleBackColor = true;
            // 
            // lblLoadOverTimeTip
            // 
            resources.ApplyResources(this.lblLoadOverTimeTip, "lblLoadOverTimeTip");
            this.lblLoadOverTimeTip.ForeColor = System.Drawing.Color.Red;
            this.lblLoadOverTimeTip.Name = "lblLoadOverTimeTip";
            // 
            // ckbSelectAllNodes
            // 
            resources.ApplyResources(this.ckbSelectAllNodes, "ckbSelectAllNodes");
            this.ckbSelectAllNodes.Checked = true;
            this.ckbSelectAllNodes.CheckState = System.Windows.Forms.CheckState.Checked;
            this.ckbSelectAllNodes.Name = "ckbSelectAllNodes";
            this.ckbSelectAllNodes.UseVisualStyleBackColor = true;
            this.ckbSelectAllNodes.CheckedChanged += new System.EventHandler(this.ckbSelectAllNodes_CheckedChanged);
            // 
            // txtFind
            // 
            this.txtFind.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtFind.ForeColor = System.Drawing.Color.Gray;
            resources.ApplyResources(this.txtFind, "txtFind");
            this.txtFind.Name = "txtFind";
            this.txtFind.MouseClick += new System.Windows.Forms.MouseEventHandler(this.txtFind_MouseClick);
            this.txtFind.TextChanged += new System.EventHandler(this.txtFind_TextChanged);
            this.txtFind.MouseLeave += new System.EventHandler(this.txtFind_MouseLeave);
            // 
            // treeViewExcels
            // 
            this.treeViewExcels.CheckBoxes = true;
            this.treeViewExcels.HideSelection = false;
            this.treeViewExcels.HotTracking = true;
            resources.ApplyResources(this.treeViewExcels, "treeViewExcels");
            this.treeViewExcels.Name = "treeViewExcels";
            this.treeViewExcels.ShowNodeToolTips = true;
            this.treeViewExcels.AfterCheck += new System.Windows.Forms.TreeViewEventHandler(this.treeViewExcels_AfterCheck);
            this.treeViewExcels.NodeMouseClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.treeViewExcels_NodeMouseClick);
            this.treeViewExcels.KeyDown += new System.Windows.Forms.KeyEventHandler(this.treeViewExcels_KeyDown);
            // 
            // tabPageExport
            // 
            this.tabPageExport.Controls.Add(this.txtFindExportTable);
            this.tabPageExport.Controls.Add(this.ListviewTableNames);
            this.tabPageExport.Controls.Add(this.tabControlDbData);
            this.tabPageExport.Controls.Add(this.btnExpertExcel);
            resources.ApplyResources(this.tabPageExport, "tabPageExport");
            this.tabPageExport.Name = "tabPageExport";
            this.tabPageExport.UseVisualStyleBackColor = true;
            // 
            // txtFindExportTable
            // 
            this.txtFindExportTable.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtFindExportTable.ForeColor = System.Drawing.Color.Gray;
            resources.ApplyResources(this.txtFindExportTable, "txtFindExportTable");
            this.txtFindExportTable.Name = "txtFindExportTable";
            this.txtFindExportTable.MouseClick += new System.Windows.Forms.MouseEventHandler(this.txtFindExportTable_MouseClick);
            this.txtFindExportTable.TextChanged += new System.EventHandler(this.txtFindExportTable_TextChanged);
            this.txtFindExportTable.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtFindExportTable_KeyDown);
            this.txtFindExportTable.MouseLeave += new System.EventHandler(this.txtFindExportTable_MouseLeave);
            // 
            // ListviewTableNames
            // 
            this.ListviewTableNames.BackColor = System.Drawing.Color.White;
            this.ListviewTableNames.CheckBoxes = true;
            this.ListviewTableNames.HideSelection = false;
            resources.ApplyResources(this.ListviewTableNames, "ListviewTableNames");
            this.ListviewTableNames.Name = "ListviewTableNames";
            this.ListviewTableNames.ShowItemToolTips = true;
            this.ListviewTableNames.UseCompatibleStateImageBehavior = false;
            this.ListviewTableNames.View = System.Windows.Forms.View.List;
            this.ListviewTableNames.ItemChecked += new System.Windows.Forms.ItemCheckedEventHandler(this.ListviewTableNames_ItemChecked);
            this.ListviewTableNames.MouseClick += new System.Windows.Forms.MouseEventHandler(this.ListviewTableNames_MouseClick);
            // 
            // tabControlDbData
            // 
            this.tabControlDbData.AllowDrop = true;
            this.tabControlDbData.Cursor = System.Windows.Forms.Cursors.Default;
            resources.ApplyResources(this.tabControlDbData, "tabControlDbData");
            this.tabControlDbData.Name = "tabControlDbData";
            this.tabControlDbData.SelectedIndex = 0;
            // 
            // tabPageLog
            // 
            this.tabPageLog.Controls.Add(this.rtxtLog);
            resources.ApplyResources(this.tabPageLog, "tabPageLog");
            this.tabPageLog.Name = "tabPageLog";
            this.tabPageLog.UseVisualStyleBackColor = true;
            // 
            // rtxtLog
            // 
            this.rtxtLog.ContextMenuStrip = this.cmsDeleteLog;
            resources.ApplyResources(this.rtxtLog, "rtxtLog");
            this.rtxtLog.Name = "rtxtLog";
            // 
            // cmsDeleteLog
            // 
            this.cmsDeleteLog.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItemClear});
            this.cmsDeleteLog.Name = "contextMenuStrip1";
            this.cmsDeleteLog.ShowImageMargin = false;
            resources.ApplyResources(this.cmsDeleteLog, "cmsDeleteLog");
            this.cmsDeleteLog.ItemClicked += new System.Windows.Forms.ToolStripItemClickedEventHandler(this.cmsDeleteLog_ItemClicked);
            // 
            // toolStripMenuItemClear
            // 
            this.toolStripMenuItemClear.Name = "toolStripMenuItemClear";
            resources.ApplyResources(this.toolStripMenuItemClear, "toolStripMenuItemClear");
            // 
            // tabPageSQL
            // 
            this.tabPageSQL.Controls.Add(this.rtxtSQL);
            resources.ApplyResources(this.tabPageSQL, "tabPageSQL");
            this.tabPageSQL.Name = "tabPageSQL";
            this.tabPageSQL.UseVisualStyleBackColor = true;
            // 
            // rtxtSQL
            // 
            resources.ApplyResources(this.rtxtSQL, "rtxtSQL");
            this.rtxtSQL.Name = "rtxtSQL";
            // 
            // label4
            // 
            resources.ApplyResources(this.label4, "label4");
            this.label4.Name = "label4";
            // 
            // txtConnetingString
            // 
            this.txtConnetingString.ForeColor = System.Drawing.Color.Black;
            resources.ApplyResources(this.txtConnetingString, "txtConnetingString");
            this.txtConnetingString.Name = "txtConnetingString";
            this.txtConnetingString.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtConnetingString_KeyDown);
            // 
            // lblDbAccess
            // 
            resources.ApplyResources(this.lblDbAccess, "lblDbAccess");
            this.lblDbAccess.ForeColor = System.Drawing.Color.Red;
            this.lblDbAccess.Name = "lblDbAccess";
            // 
            // btnTestDbConnection
            // 
            resources.ApplyResources(this.btnTestDbConnection, "btnTestDbConnection");
            this.btnTestDbConnection.Name = "btnTestDbConnection";
            this.btnTestDbConnection.UseVisualStyleBackColor = true;
            this.btnTestDbConnection.Click += new System.EventHandler(this.btnTestDbConnection_Click);
            // 
            // cmsSheetNode
            // 
            this.cmsSheetNode.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItemCopy,
            this.toolStripSeparator1,
            this.toolStripMenuMapping,
            this.toolStripSeparator2,
            this.toolStripTextBoxComment});
            this.cmsSheetNode.Name = "contextMenuStrip1";
            this.cmsSheetNode.ShowImageMargin = false;
            resources.ApplyResources(this.cmsSheetNode, "cmsSheetNode");
            this.cmsSheetNode.Opened += new System.EventHandler(this.cmsSheetNode_Opened);
            this.cmsSheetNode.ItemClicked += new System.Windows.Forms.ToolStripItemClickedEventHandler(this.cmsSheetNode_ItemClicked);
            // 
            // toolStripMenuItemCopy
            // 
            this.toolStripMenuItemCopy.Name = "toolStripMenuItemCopy";
            resources.ApplyResources(this.toolStripMenuItemCopy, "toolStripMenuItemCopy");
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            resources.ApplyResources(this.toolStripSeparator1, "toolStripSeparator1");
            // 
            // toolStripMenuMapping
            // 
            this.toolStripMenuMapping.Name = "toolStripMenuMapping";
            resources.ApplyResources(this.toolStripMenuMapping, "toolStripMenuMapping");
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            resources.ApplyResources(this.toolStripSeparator2, "toolStripSeparator2");
            // 
            // toolStripTextBoxComment
            // 
            this.toolStripTextBoxComment.Name = "toolStripTextBoxComment";
            resources.ApplyResources(this.toolStripTextBoxComment, "toolStripTextBoxComment");
            this.toolStripTextBoxComment.KeyDown += new System.Windows.Forms.KeyEventHandler(this.toolStripTextBoxComment_KeyDown);
            this.toolStripTextBoxComment.Click += new System.EventHandler(this.toolStripTextBoxComment_Click);
            // 
            // cmsOpenRoot
            // 
            this.cmsOpenRoot.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItemOpenRoot,
            this.toolStripSyncConfig,
            this.toolStripUpdateTime,
            this.toolStripReboot});
            this.cmsOpenRoot.Name = "cmsOpenRoot";
            this.cmsOpenRoot.ShowImageMargin = false;
            resources.ApplyResources(this.cmsOpenRoot, "cmsOpenRoot");
            // 
            // toolStripMenuItemOpenRoot
            // 
            this.toolStripMenuItemOpenRoot.Name = "toolStripMenuItemOpenRoot";
            resources.ApplyResources(this.toolStripMenuItemOpenRoot, "toolStripMenuItemOpenRoot");
            this.toolStripMenuItemOpenRoot.Click += new System.EventHandler(this.toolStripMenuItemOpenRoot_Click);
            // 
            // toolStripSyncConfig
            // 
            this.toolStripSyncConfig.Name = "toolStripSyncConfig";
            resources.ApplyResources(this.toolStripSyncConfig, "toolStripSyncConfig");
            this.toolStripSyncConfig.Click += new System.EventHandler(this.toolStripSyncConfig_Click);
            // 
            // toolStripUpdateTime
            // 
            this.toolStripUpdateTime.Name = "toolStripUpdateTime";
            resources.ApplyResources(this.toolStripUpdateTime, "toolStripUpdateTime");
            this.toolStripUpdateTime.Click += new System.EventHandler(this.toolStripUpdateTime_Click);
            // 
            // toolStripReboot
            // 
            this.toolStripReboot.Name = "toolStripReboot";
            resources.ApplyResources(this.toolStripReboot, "toolStripReboot");
            this.toolStripReboot.Click += new System.EventHandler(this.toolStripReboot_Click);
            // 
            // Main
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ContextMenuStrip = this.cmsOpenRoot;
            this.Controls.Add(this.btnTestDbConnection);
            this.Controls.Add(this.lblDbAccess);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.txtConnetingString);
            this.Controls.Add(this.tabControlMain);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "Main";
            this.tabControlMain.ResumeLayout(false);
            this.tabPageImport.ResumeLayout(false);
            this.tabPageImport.PerformLayout();
            this.tabPageExport.ResumeLayout(false);
            this.tabPageExport.PerformLayout();
            this.tabPageLog.ResumeLayout(false);
            this.cmsDeleteLog.ResumeLayout(false);
            this.tabPageSQL.ResumeLayout(false);
            this.cmsSheetNode.ResumeLayout(false);
            this.cmsSheetNode.PerformLayout();
            this.cmsOpenRoot.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnExpertExcel;
        private System.Windows.Forms.TabControl tabControlSheetInfo;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtSelectExcel;
        private System.Windows.Forms.Button btnImportBatchSheets;
        private System.Windows.Forms.ProgressBar pgbBatchImport;
        private System.Windows.Forms.Label lblImportTableName;
        private System.Windows.Forms.Label lblSheetInfo;
        private System.Windows.Forms.TabControl tabControlMain;
        private System.Windows.Forms.TabPage tabPageImport;
        private System.Windows.Forms.TabPage tabPageExport;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtConnetingString;
        private System.Windows.Forms.Label lblDbAccess;
        private System.Windows.Forms.TreeView treeViewExcels;
        private System.Windows.Forms.Button btnTestDbConnection;
        private System.Windows.Forms.TabPage tabPageLog;
        private System.Windows.Forms.RichTextBox rtxtLog;
        private System.Windows.Forms.ContextMenuStrip cmsDeleteLog;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItemClear;
        private System.Windows.Forms.TextBox txtFind;
        private System.Windows.Forms.CheckBox ckbSelectAllNodes;
        private System.Windows.Forms.ContextMenuStrip cmsSheetNode;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItemCopy;
        private System.Windows.Forms.TabPage tabPageSQL;
        private System.Windows.Forms.RichTextBox rtxtSQL;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuMapping;
        private System.Windows.Forms.TabControl tabControlDbData;
        private System.Windows.Forms.ListView ListviewTableNames;
        private System.Windows.Forms.TextBox txtFindExportTable;
        private System.Windows.Forms.Label lblLoadOverTimeTip;
        private System.Windows.Forms.ContextMenuStrip cmsOpenRoot;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItemOpenRoot;
        private System.Windows.Forms.ToolStripMenuItem toolStripSyncConfig;
        private System.Windows.Forms.ToolStripMenuItem toolStripUpdateTime;
        private System.Windows.Forms.ToolStripMenuItem toolStripReboot;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripTextBox toolStripTextBoxComment;
    }
}