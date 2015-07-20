// ****************************************
// FileName:Main.cs
// Description:Excel工具主界面类
// Tables:Many
// Author:Gavin && Burney
// Create Date:2014-06-01
// Revision History:
// ****************************************

using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.IO;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Windows.Forms;

namespace ExcelTool
{
    using BLL;
    using Model;
    using Utils.Configuration;
    using Utils.Excel;
    using Utils.Xml;

    public partial class Main : Form
    {
        #region 字段与属性

        private static readonly Object lockObj = new Object();

        //缓存-所选Excel表单信息 (Key:文件名; Value:文件全路径名)
        internal Dictionary<String, String> mExcelFileInfos;

        //缓存-已读取的Excel表单
        internal DataSet mAllTables = new DataSet();

        //缓存-已读取过的Excel对象
        internal Dictionary<String, MoqikakaExcel> mExcels = new Dictionary<String, MoqikakaExcel>();

        //缓存-已存在映射关系的Excel表单名列表
        private List<String> mMappingedSheetList;

        //日志文件路径
        private String mLogFilePath = CommonBLL.GetCurrentLogFilePath();

        //当前操作sql文件路径
        private String CurrentSqlFilePath;

        /// <summary>
        /// 当前数据库所有表名
        /// </summary>
        private List<String> AllDbTableNames { get; set; }

        /// <summary>
        /// 页面显示表单数据的最大数量
        /// </summary>
        private Int32 ShowSheetDataCount
        {
            get
            {
                Int32 num = -1;
                Int32.TryParse(ConfigurationHelper.AppSettings["ShowSheetDataCount"], out num);
                return num;
            }
        }

        #endregion

        #region 页面初始化

        /// <summary>
        /// 构造函数,初始化控件信息
        /// </summary>
        public Main()
        {
            InitializeComponent();

            //加载数据库连接字符串
            LoadConnectionString();

            //设置选择TextBox的值
            InitSelectText(ConstantText.ClickToChoose);

            //加载已映射的表单列表
            InitMappingSheet();

            //测试连接字符串状态
            btnTestDbConnection_Click(null, null);
        }

        /// <summary>
        /// 加载数据库连接字符串
        /// </summary>
        private void LoadConnectionString()
        {
            //读取app.config数据
            txtConnetingString.Text = ConfigurationHelper.ConnectionString.Value;
        }

        /// <summary>
        /// 设置选择TextBox的值
        /// </summary>
        /// <param name="value">设定值</param>
        private void InitSelectText(String value)
        {
            //设置文本框值和状态
            txtSelectExcel.Text = value;
            txtSelectExcel.SelectionStart = txtSelectExcel.Text.Length;
            txtConnetingString.SelectionStart = txtConnetingString.Text.Length;
        }

        /// <summary>
        /// 初始化加载已存在映射的表单列表
        /// </summary>
        private void InitMappingSheet()
        {
            XMLHelper xmlHelper = new XMLHelper("TableMapping.xml");
            mMappingedSheetList = xmlHelper.LoadMappingedSheetNameList();
        }

        /// <summary>
        /// 初始化数据库数据
        /// </summary>
        private void InitDbTables()
        {
            if (ExcelBLL.IsDataBaseAccess())
            {
                AllDbTableNames = ExcelBLL.GetTableNameList();

                RefreshTableNames();
            }
        }

        /// <summary>
        /// 刷新数据库表名
        /// </summary>
        private void RefreshTableNames()
        {
            //刷新导入界面
            foreach (TreeNode item in treeViewExcels.Nodes)
            {
                foreach (TreeNode sheet in item.Nodes)
                {
                    //存在映射关系
                    if (mMappingedSheetList.Contains(sheet.Text.ToUpper()))
                    {
                        sheet.ForeColor = Color.Red;
                        sheet.ToolTipText = ConstantText.NoMappingRelation;
                        continue;
                    }

                    //数据库存在该表
                    if (!AllDbTableNames.Contains(sheet.Text.ToLower()))
                    {
                        sheet.ForeColor = Color.Gray;
                        sheet.ToolTipText = ConstantText.TableNotExist; //"当前数据库不存在该表";
                        continue;
                    }

                    sheet.ForeColor = Color.Black;
                    sheet.ToolTipText = ConstantText.ClickToCopy; //"点击复制表名";
                }
            }

            //刷新导出界面
            ListviewTableNames.Clear();
            tabControlDbData.TabPages.Clear();

            //将数据库表格列表绑定到界面上
            for (Int32 i = 0; i < AllDbTableNames.Count; i++)
            {
                //字段名todo
                ListViewItem item = new ListViewItem(AllDbTableNames[i])
                {
                    ToolTipText = "单击查看该表数据"
                };

                ListviewTableNames.Items.Add(item);
            }
        }

        #endregion

        #region 01 导入页面

        /// <summary>
        /// Excel选择文本框 鼠标点击事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void txtSelectExcel_MouseClick(Object sender, MouseEventArgs e)
        {
            #region 选择文件对话框

            //构造弹出对话框
            OpenFileDialog diog = new OpenFileDialog
            {
                Multiselect = true,
                Title = CommonBLL.GetDialogTitle(),
                Filter = ConstantText.ExcelFileFilter //@"Excel文档|*.xls;*.xlsx",
            };

            //未选择直接返回
            if (diog.ShowDialog() != DialogResult.OK || diog.FileNames.Length == 0)
                return;

            #endregion

            #region 重置数据

            //清空缓存
            mAllTables.Clear();
            treeViewExcels.Nodes.Clear();
            tabControlSheetInfo.TabPages.Clear();
            mExcelFileInfos = new Dictionary<String, String>();

            #endregion

            #region 收集所选文件

            //收集所选Excel信息
            for (int i = 0; i < diog.SafeFileNames.Length; i++)
            {
                if (diog.SafeFileNames[i].StartsWith("~")) continue;

                mExcelFileInfos.Add(diog.SafeFileNames[i], diog.FileNames[i]);
            }

            #endregion

            #region 异步加载

            //加载中
            btnImportBatchSheets.Text = "加载中";
            btnImportBatchSheets.Enabled = false;

            //异步加载Excel表单节点
            Action<Dictionary<String, String>> action = new Action<Dictionary<String, String>>(LoadExcelSheets);
            action.BeginInvoke(mExcelFileInfos, null, null);

            #endregion

            #region 设置控件属性

            //显示友好提示信息
            lblLoadOverTimeTip.Visible = true;
            InitSelectText(String.Join(";  ", diog.SafeFileNames));

            #endregion
        }

        /// <summary>
        /// 批量导入
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnImportBatchSheets_Click(Object sender, EventArgs e)
        {
            //防止连续点击
            btnImportBatchSheets.Enabled = false;

            //重置忽略表单
            IgnoreSheetsBLL.Reset();

            //是否需要创建表
            Boolean needCreateTable = ExcelBLL.IfCreateTable();

            //存放Excel名称和所选表单信息
            Dictionary<String, String> selectedSheets = new Dictionary<String, String>();

            //遍历添加所有勾选的表单和Excel路径
            foreach (TreeNode firstLevelNode in treeViewExcels.Nodes)
            {
                foreach (TreeNode secondLevelNode in firstLevelNode.Nodes)
                {
                    if (secondLevelNode.Checked)
                    {
                        //不需要创建表&&数据库表不存在时,不导入
                        if (!needCreateTable && secondLevelNode.ForeColor == Color.Gray)
                            continue;

                        //存在表字段名行时,不导入
                        if (ExcelBLL.IsUselessSheet(secondLevelNode.Text))
                            continue;

                        //忽略表单
                        if (IgnoreSheetsBLL.IsIgnoreSheet(secondLevelNode.Text))
                            continue;

                        //有重复表单名时,不导入
                        if (selectedSheets.ContainsKey(secondLevelNode.Text))
                            continue;

                        selectedSheets.Add(secondLevelNode.Text, mExcelFileInfos[firstLevelNode.Text]);
                    }
                }
            }

            //没有可导入的数据
            if (selectedSheets.Count == 0)
            {
                MessageBox.Show("没有可导入的数据, 或配置中CreateTable == false!");
                btnImportBatchSheets.Enabled = true;
                return;
            }

            //检测数据库连接字符串是否可用
            if (!ExcelBLL.IsDataBaseAccess())
            {
                MessageBox.Show(@"请检查该数据库连接字符串是否正确!");
                btnImportBatchSheets.Enabled = true;
                return;
            }

            //设置当前操作SQL存放文件(每导入单独存放)
            CurrentSqlFilePath = CommonBLL.GetCurrentSqlFilePath();

            //异步批量导入
            Func<Dictionary<String, String>, Dictionary<String, Int32>> func = new Func<Dictionary<String, String>, Dictionary<String, Int32>>(Import);
            func.BeginInvoke(selectedSheets, ImportCallBack, null);
        }

        /// <summary>
        /// 表单标签页改变事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tabControlSheetInfo_SelectedIndexChanged(Object sender, EventArgs e)
        {
            if (tabControlSheetInfo.TabPages.Count == 0)
            {
                lblSheetInfo.Text = String.Empty;
                return;
            }

            String sheetName = tabControlSheetInfo.SelectedTab.Text;

            Int32 sheetRowCount = !mAllTables.Tables.Contains(sheetName) ? 0 : mAllTables.Tables[sheetName].Rows.Count - MoqikakaExcelSettings.SpecialRowList.Count + 1;

            lblSheetInfo.Text = String.Format("表单:{0}  数据总数:{1}", sheetName, sheetRowCount);
        }

        /// <summary>
        /// tv_Excels节点点击事件,显示表单数据 或者 显示菜单
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void treeViewExcels_NodeMouseClick(Object sender, TreeNodeMouseClickEventArgs e)
        {
            //所点击的结点名称
            String clickNodeName = e.Node.Text;

            if (e.Button == MouseButtons.Right)
            {
                treeViewExcels.SelectedNode = e.Node;
                return;
            }

            //不是有效的点击
            if (NotValidClick(e)) return;

            if (e.Node.Level == 0)      //第一阶菜单,默认展开当前Excel所有表单
            {
                ShowExcelSheets(clickNodeName);
            }
            else if (e.Node.Level == 1) //第二级菜单,显示单个sheet
            {
                ShowSheet(clickNodeName, e.Node.Parent.Text);
            }
        }

        /// <summary>
        /// tv_Excels节点选中事件,选中第一阶点时,全选/全取消 子节点
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void treeViewExcels_AfterCheck(Object sender, TreeViewEventArgs e)
        {
            //第一阶节点,全选子节点
            if (e.Node.Level == 0)
            {
                foreach (TreeNode node in e.Node.Nodes)
                {
                    node.Checked = e.Node.Checked;
                }
            }
        }

        /// <summary>
        /// 快捷键
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void treeViewExcels_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.C)
            {
                cmsSheetNode.Show(Cursor.Position);
                cmsSheetNode.Focus();
                toolStripTextBoxComment.Focus();
            }
        }

        /// <summary>
        /// 测试按钮点击事件 测试连接字符串是否可连接到数据库
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnTestDbConnection_Click(Object sender, EventArgs e)
        {
            String connectionString = txtConnetingString.Text.Trim();

            if (String.IsNullOrEmpty(connectionString)) return;

            //将ConnectionString写入app.config配置文件
            ConfigurationHelper.ConnectionString.Value = connectionString;

            //设置数据连接字符串
            ExcelBLL.SetDbConnection(connectionString);

            //判断数据库连接是否可用
            if (!ExcelBLL.IsDataBaseAccess())
            {
                this.btnImportBatchSheets.Enabled = false;
                lblDbAccess.Image = Image.FromFile(Environment.CurrentDirectory + "\\Icon\\0.jpg"); //连接不成功
            }
            else
            {
                //测试成功后方能导入
                this.btnImportBatchSheets.Enabled = true;
                this.btnImportBatchSheets.Text = ConstantText.Import; //"导入";
                lblDbAccess.Image = Image.FromFile(Environment.CurrentDirectory + "\\Icon\\1.jpg");  //连接成功

                //重新加载导出表数据
                InitDbTables();
            }
        }

        /// <summary>
        /// 连接字符串输入框键盘按下事件,数据库连接字符串编辑时,按回车键测试连接
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void txtConnetingString_KeyDown(Object sender, KeyEventArgs e)
        {
            //回车键,触发测试按钮点击事件,检测连接字符串是否可用
            if (e.KeyCode == Keys.Enter)
            {
                btnTestDbConnection_Click(null, null);
                return;
            }

            this.btnImportBatchSheets.Enabled = false;
            this.btnImportBatchSheets.Text = ConstantText.TestPlease;
        }

        /// <summary>
        /// 查找输入框文本改变事件,动态查找匹配结点
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void txtFind_TextChanged(Object sender, EventArgs e)
        {
            String input = txtFind.Text.Trim();

            //匹配所有表单名
            foreach (TreeNode node in treeViewExcels.Nodes)
            {
                node.Collapse();

                foreach (TreeNode item in node.Nodes)
                {
                    if (!ContainsText(item.Text, input))
                    {
                        item.BackColor = Color.White;
                        continue;
                    }

                    node.Expand();
                    item.BackColor = Color.Brown;
                }
            }
        }

        /// <summary>
        ///全选的CheckBox 的Check改变事件,TreeView结点全选
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ckbSelectAllNodes_CheckedChanged(Object sender, EventArgs e)
        {
            //全选 或 全部取消 第一阶节点 (会联动,触发treeViewExcels_AfterCheck)
            foreach (TreeNode node in treeViewExcels.Nodes)
            {
                node.Checked = ckbSelectAllNodes.Checked;
            }
        }

        /// <summary>
        /// 点击查找框
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void txtFind_MouseClick(Object sender, MouseEventArgs e)
        {
            if (txtFind.Text == ConstantText.Find)
            {
                txtFind.Text = String.Empty;
            }
        }

        /// <summary>
        /// 鼠标离开查找输入框
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void txtFind_MouseLeave(Object sender, EventArgs e)
        {
            if (txtFind.Text == String.Empty)
            {
                txtFind.Text = ConstantText.Find;
            }
        }

        /// <summary>
        /// 节点右键菜单打开
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cmsSheetNode_Opened(object sender, EventArgs e)
        {
            //获取所点击的结点名称
            String sheetName = treeViewExcels.SelectedNode.Text;

            //是否显示表名备注
            Boolean showTableNameComment = treeViewExcels.SelectedNode.ForeColor == Color.Gray;

            toolStripSeparator2.Visible = showTableNameComment;
            toolStripTextBoxComment.Visible = showTableNameComment;

            if (showTableNameComment)
            {
                toolStripTextBoxComment.Text = ConstantText.AddComment;
                toolStripTextBoxComment.ToolTipText = ConstantText.AddCommentTips;

                String comment = ExcelBLL.GetTableComment(sheetName);
                if (comment != null)
                {
                    toolStripTextBoxComment.Text = comment;
                }
            }
        }

        /// <summary>
        /// toolStripTextBoxComment,键盘按下事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripTextBoxComment_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode != Keys.Enter) return;

            if (toolStripTextBoxComment.Text == ConstantText.AddComment) return;

            //保存备注信息
            ExcelBLL.AddTableComment(treeViewExcels.SelectedNode.Text, toolStripTextBoxComment.Text);

            cmsSheetNode.Close();
        }

        /// <summary>
        /// toolStripTextBoxComment,单击事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripTextBoxComment_Click(object sender, EventArgs e)
        {
            toolStripTextBoxComment.Text = String.Empty;
        }

        /// <summary>
        /// TreeNode结点,右键菜单选中事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cmsSheetNode_ItemClicked(Object sender, ToolStripItemClickedEventArgs e)
        {
            //获取所点击的结点名称
            String sheetName = treeViewExcels.SelectedNode.Text;

            //复制表名
            if (e.ClickedItem.Text.Contains(ConstantText.CopySheetName))
            {
                Clipboard.SetText(sheetName);
                return;
            }

            //映射该表
            if (e.ClickedItem.Text.Contains(ConstantText.MappingTable))
            {
                //实例化弹出映射窗体
                TableMapping tableMapping = new TableMapping(sheetName, mExcelFileInfos[treeViewExcels.SelectedNode.Parent.Text], this);
                tableMapping.ShowDialog(this);

                //改变所选节点,映射状态
                XMLHelper xmlHelper = new XMLHelper("TableMapping.xml");

                if (xmlHelper.GetTableNameMapping(sheetName) != null)
                {
                    treeViewExcels.SelectedNode.ForeColor = Color.Red;
                    treeViewExcels.SelectedNode.ToolTipText = ConstantText.MappingExist;
                }
                else
                {
                    treeViewExcels.SelectedNode.ForeColor = Color.Black;
                    treeViewExcels.SelectedNode.ToolTipText = ConstantText.ClickToCopy;
                }
            }
        }

        #endregion

        #region 02 导出页面

        /// <summary>
        /// 导出
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnExportExcel_Click(Object sender, EventArgs e)
        {
            if (ListviewTableNames.CheckedItems.Count == 0) return;

            //选择存放路径
            FolderBrowserDialog fbd = new FolderBrowserDialog
            {
                ShowNewFolderButton = true,
                Description = ConstantText.ExportDescription
            };

            //上次导出Excel文档存放文件路径
            String exportExcelFilePath = CommonBLL.GetStoredFolder();

            //如果存在上次导出所选路径
            if (!String.IsNullOrEmpty(exportExcelFilePath)) fbd.SelectedPath = exportExcelFilePath;

            if (fbd.ShowDialog(this) != DialogResult.OK) return;

            btnExpertExcel.Enabled = false;
            String selectedPath = fbd.SelectedPath;

            //保存所选路径
            CommonBLL.StoreExportFolder(selectedPath);

            Int32 successCount = 0;  //成功导出个数

            foreach (ListViewItem item in ListviewTableNames.CheckedItems)
            {
                String tableName = item.Text.Trim();
                String filePath = Path.Combine(selectedPath, tableName + ".xlsx");

                //获取导出数据
                DataTable dt = ExcelBLL.GetTableData(tableName);
                dt.TableName = tableName;

                //导出Excel
                MoqikakaExcel.Write(dt, filePath, ExcelBLL.GetComments(tableName));

                successCount++;
            }

            btnExpertExcel.Enabled = true;

            DialogResult result = MessageBox.Show(ConstantText.ShowFolder, String.Format(ConstantText.ExportSuccessTips, successCount), MessageBoxButtons.OKCancel);
            if (result == DialogResult.OK)
            {
                Process.Start(selectedPath);
            }
        }

        /// <summary>
        /// 导出表名鼠标点击事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ListviewTableNames_MouseClick(Object sender, MouseEventArgs e)
        {
            //Check时,不绑定表数据
            if (e.Location.X <= 15) return;

            //获取所点击的表名
            String tableName = ListviewTableNames.FocusedItem.Text;

            //若存在该表标签,选中该标签
            foreach (TabPage page in tabControlDbData.TabPages)
            {
                //若TablePages存在该表单,则选中该TabPage
                if (page.Text == tableName)
                {
                    tabControlDbData.SelectedTab = page;
                    return;
                }
            }

            ShowTableData(tableName);
        }

        /// <summary>
        /// 数据库表CheckBox选中后触发事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ListviewTableNames_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            //Check后加粗节点字体
            e.Item.Font = e.Item.Checked ? new Font(DefaultFont.FontFamily.Name, 9, FontStyle.Bold) : DefaultFont;
        }

        /// <summary>
        /// 导出页面-查找输入框文本改变事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void txtFindExportTable_TextChanged(Object sender, EventArgs e)
        {
            String tableName = txtFindExportTable.Text;

            Boolean locatedFirstPostion = false;  //是否 已定位到第一满足项 (防止焦点不断移动)

            //匹配所有节点
            foreach (ListViewItem item in ListviewTableNames.Items)
            {
                if (ContainsText(item.Text, tableName))
                {
                    if (!locatedFirstPostion) //仅定位一次
                    {
                        // 滚动滑动条使该项可见
                        ListviewTableNames.EnsureVisible(item.Index);
                        locatedFirstPostion = true;
                    }

                    item.Checked = true;
                }
                else
                {
                    //恢复样式
                    item.BackColor = Color.White;
                    item.Checked = false;
                    item.Selected = false;
                }
            }
        }

        /// <summary>
        /// 导出页面-查找输入框鼠标点击事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void txtFindExportTable_MouseClick(Object sender, MouseEventArgs e)
        {
            if (txtFindExportTable.Text == ConstantText.Find)
            {
                txtFindExportTable.Text = String.Empty;
            }
        }

        /// <summary>
        /// 导出页面-查找输入框鼠标离开事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void txtFindExportTable_MouseLeave(Object sender, EventArgs e)
        {
            if (txtFindExportTable.Text == String.Empty)
            {
                txtFindExportTable.Text = ConstantText.Find;
            }
        }

        /// <summary>
        /// 搜索框Enter确认所选表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void txtFindExportTable_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Enter)
            {
                //清空当前TabPage页
                tabControlDbData.TabPages.Clear();

                //添加所有匹配项的标签页
                foreach (ListViewItem item in ListviewTableNames.Items)
                {
                    if (!item.Checked) continue;

                    AddTabPage(item.Text);
                }
            }
        }

        #endregion

        #region 03 其他页面

        /// <summary>
        /// 打开目录
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripMenuItemOpenRoot_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(Environment.CurrentDirectory);
        }

        /// <summary>
        /// 同步设置
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripSyncConfig_Click(object sender, EventArgs e)
        {
            MoqikakaExcelSettings.Init();
        }

        /// <summary>
        /// 更新Check时间
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripUpdateTime_Click(object sender, EventArgs e)
        {
            ExcelBLL.UpdateCheckInfoTime();
        }

        /// <summary>
        /// 重新启动
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripReboot_Click(object sender, EventArgs e)
        {
            String path = Path.Combine(Environment.CurrentDirectory, "ExcelTool.exe");

            if (File.Exists(path))
            {
                System.Diagnostics.Process.Start(path);
                this.Close();
            }
        }

        /// <summary>
        /// 菜单选项卡改变事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tabControlMain_SelectedIndexChanged(Object sender, EventArgs e)
        {
            String defaultText = @"没有记录";
            switch (tabControlMain.SelectedIndex)
            {
                case 2://选中日志标签页
                    rtxtLog.Text = File.Exists(mLogFilePath) ? File.ReadAllText(mLogFilePath) : defaultText;
                    if (rtxtLog.TextLength == 0) rtxtLog.Text = defaultText;
                    break;
                case 3://选中SQL标签页
                    rtxtSQL.Text = File.Exists(CurrentSqlFilePath) ? File.ReadAllText(CurrentSqlFilePath) : defaultText;
                    if (rtxtSQL.TextLength == 0) rtxtSQL.Text = defaultText;
                    break;
                default:
                    return;
            }
        }

        /// <summary>
        /// 鼠标右键,清空日志
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cmsDeleteLog_ItemClicked(Object sender, ToolStripItemClickedEventArgs e)
        {
            if (!e.ClickedItem.Text.Contains("清空")) return;

            //清空显示
            rtxtLog.Text = @"没有日志记录";

            //清空文件
            File.WriteAllText(mLogFilePath, String.Empty);
        }

        #endregion

        #region 内部调用方法

        /// <summary>
        /// 显示单个表单数据
        /// </summary>
        /// <param name="sheetName">表单名</param>
        /// <param name="excelName">所属excel名</param>
        private void ShowSheet(String sheetName, String excelName)
        {
            //如果tabControlSheetInfo中存在该标签,则选中该标签
            foreach (TabPage tab in tabControlSheetInfo.TabPages)
            {
                if (tab.Text == sheetName)
                {
                    tabControlSheetInfo.SelectedTab = tab;
                    return;
                }
            }

            //获取表单数据
            DataTable dt = GetSheetTable(sheetName, mExcelFileInfos[excelName]);
            if (dt == null) dt = new DataTable(sheetName);

            //创建新表单标签页
            CreateNewTabPage(dt);

            lblSheetInfo.Text = String.Format("表单:{0}  数据总数:{1}", sheetName, dt.Rows.Count == 0 ? 0 : dt.Rows.Count - MoqikakaExcelSettings.SpecialRowList.Count + 1);
        }

        /// <summary>
        /// 显示整个excel表单数据
        /// </summary>
        /// <param name="excelName">所属excel名</param>
        private void ShowExcelSheets(string excelName)
        {
            //删除已有标签
            tabControlSheetInfo.TabPages.Clear();

            //异步加载表单数据
            Action<String> loadSheet = new Action<String>(BindSheets);
            loadSheet.BeginInvoke(mExcelFileInfos[excelName], null, null);

        }

        /// <summary>
        /// 不是有效的点击
        /// </summary>
        /// <param name="e"></param>
        /// <returns></returns>
        private bool NotValidClick(TreeNodeMouseClickEventArgs e)
        {
            if (e.Node.Level == 0 && e.Location.X <= 35) return true;

            if (e.Node.Level == 1 && e.Location.X <= 55) return true;

            return false;
        }

        /// <summary>
        /// 是否包含文本
        /// </summary>
        /// <param name="sheetName">表单名</param>
        /// <param name="searchText">搜索文本</param>
        /// <returns>是否包含文本</returns>
        private bool ContainsText(string sheetName, string searchText)
        {
            if (String.IsNullOrEmpty(searchText)) return false;

            return sheetName.IndexOf(searchText, StringComparison.Ordinal) != -1;
        }

        /// <summary>
        /// 显示表数据
        /// </summary>
        /// <param name="tableName">表名</param>
        private void ShowTableData(string tableName)
        {
            //最多显示5个TabPage
            if (tabControlDbData.TabPages.Count > 4)
            {
                tabControlDbData.TabPages.RemoveAt(0);
            }

            AddTabPage(tableName);

            KeepShowTableColor();
        }

        /// <summary>
        /// 显示在TabPage的表名变色
        /// </summary>
        private void KeepShowTableColor()
        {
            //保留前五次所点击的项
            List<String> lastFiveClickedItem = new List<String>();
            foreach (TabPage item in tabControlDbData.TabPages)
            {
                lastFiveClickedItem.Add(item.Text);
            }

            //对保存项TabPage进行变色
            foreach (ListViewItem item in ListviewTableNames.Items)
            {
                item.BackColor = lastFiveClickedItem.Contains(item.Text) ? Color.Gainsboro : Color.White;
            }
        }

        /// <summary>
        /// 添加数据库表tab页
        /// </summary>
        /// <param name="tableName">表名</param>
        private void AddTabPage(string tableName)
        {
            TabPage newPage = new TabPage(tableName);

            //获取并绑定数据
            DataTable tableData = ExcelBLL.GetTableData(tableName);

            newPage.Controls.Add(CreatSingleGridview(tableData));

            tabControlDbData.TabPages.Add(newPage);
            tabControlDbData.SelectedTab = newPage;
        }

        /// <summary>
        /// 选中多个Excel后,采用异步加载多个Excel中的表单节点
        /// </summary>
        /// <param name="selectedExcelInfos">选择的Excel表单信息</param>
        private void LoadExcelSheets(Dictionary<String, String> selectedExcelInfos)
        {
            //todo refactoring to here

            //是否为第一个Excel节点
            Boolean firstExcelNode = true;

            //最后一个excel
            var lastExcel = selectedExcelInfos.OrderBy(p => new FileInfo(p.Value).Length).Last().Value;

            //循环处理每个Excel
            foreach (var item in selectedExcelInfos.OrderBy(p => new FileInfo(p.Value).Length))
            {
                MethodInvoker invoker = new MethodInvoker(() =>
                {
                    MoqikakaExcel excel = LoadExcel(item.Value);

                    TreeNode[] nodes = GetSheetNodesByExcelFile(excel);

                    //跨线程向treeViewExcels添加节点
                    InvokeMethod(() =>
                    {
                        //绑定TreeView
                        TreeNode listViewItem = new TreeNode(item.Key, nodes)
                        {
                            Checked = true,
                            ToolTipText = ConstantText.ShowAllSheets
                        };

                        if (firstExcelNode)
                        {
                            listViewItem.ExpandAll();
                            firstExcelNode = false;
                        }

                        treeViewExcels.Nodes.Add(listViewItem);
                    });
                });

                invoker.BeginInvoke(p =>
                {
                    if (item.Value != lastExcel) return;

                    InvokeMethod(() =>
                    {
                        btnImportBatchSheets.Text = "导入";
                        btnImportBatchSheets.Enabled = true;
                    });
                }, null);
            }
        }

        /// <summary>
        /// 根据Excel名,构造表单的节点数组
        /// </summary>
        /// <param name="excelFilePath">Excel文件路径</param>
        /// <returns></returns>
        private TreeNode[] GetSheetNodesByExcelFile(MoqikakaExcel excel)
        {
            //获取Excel所有表单列表
            List<String> sheetList = excel.SheetNameList;

            //存放表单节点的数组
            TreeNode[] nodeArray = new TreeNode[sheetList.Count];

            //根据每个Excel的表单名,创建对应的节点
            for (Int32 i = 0; i < sheetList.Count; i++)
            {
                TreeNode node = new TreeNode(sheetList[i])
                {
                    ContextMenuStrip = cmsSheetNode,
                    Checked = true,
                    ToolTipText = "右击可复制表名"
                };

                //对已存在映射关系的节点
                if (mMappingedSheetList.Contains(sheetList[i].ToUpper()))
                {
                    node.ForeColor = Color.Red;
                    node.ToolTipText = "该表存在映射关系";
                }

                if (!AllDbTableNames.Contains(sheetList[i].ToLower()))
                {
                    node.ForeColor = Color.Gray;
                    node.ToolTipText = "当前数据库不存在该表";
                }

                //添加到对应节点
                nodeArray[i] = node;
            }

            return nodeArray;
        }

        /// <summary>
        /// 绑定一个Excel的所有表单信息 (异步)
        /// </summary>
        /// <param name="path">文件路径</param>
        private void BindSheets(String path)
        {
            //加载excel对象
            MoqikakaExcel excel = LoadExcel(path);

            #region 创建表单标签页

            TabPage[] pages = new TabPage[excel.NumberOfSheets];

            //初始化TabPage标签
            for (Int32 i = 0; i < excel.NumberOfSheets; i++)
            {
                TabPage page = new TabPage(excel.GetSheetName(i));
                pages[i] = page;
            }

            //跨线程改变主线程控件
            tabControlSheetInfo.Invoke(new MethodInvoker(() => tabControlSheetInfo.TabPages.AddRange(pages)));

            #endregion

            //附加数据到每个TabPage页面
            AppendDataToTabPages(excel);
        }

        /// <summary>
        /// 为每个TabPage附加数据
        /// </summary>
        /// <param name="excel">Excel对象</param>
        private void AppendDataToTabPages(MoqikakaExcel excel)
        {
            DataTable table = null;
            String sheetName;

            //循环读取每个Excel表单的数据
            for (Int32 i = 0; i < excel.NumberOfSheets; i++)
            {
                sheetName = excel.GetSheetName(i);

                table = GetSheetTable(sheetName, excel);

                //通知主线程控件,数据已准备好
                tabControlSheetInfo.Invoke(new Action(() =>
                {
                    if (tabControlSheetInfo.TabPages.Count > i)
                    {
                        tabControlSheetInfo.TabPages[i].Controls.Add(CreatSingleGridview(table));
                    }
                }));
            }
        }

        /// <summary>
        /// 创建单个DataGridView,并绑定数据源
        /// </summary>
        /// <param name="dt">数据源</param>
        /// <returns>绑定好数据源的DataGridView</returns>
        private DataGridView CreatSingleGridview(DataTable dt)
        {
            //创建新的DataGridView,并设置其样式和绑定数据源
            DataGridView dgv = new DataGridView();
            dgv.BackgroundColor = Color.White;
            dgv.DefaultCellStyle.BackColor = Color.White;
            dgv.DefaultCellStyle.ForeColor = Color.Black;
            dgv.DefaultCellStyle.SelectionBackColor = Color.Goldenrod;
            dgv.RowHeadersWidth = 25;
            dgv.AllowUserToAddRows = false;
            dgv.ReadOnly = true;
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dgv.Dock = DockStyle.Fill;
            dgv.ScrollBars = ScrollBars.Both;
            dgv.DataSource = GetShowTable(dt);

            return dgv;
        }

        /// <summary>
        /// 获取显示的表数据
        /// </summary>
        /// <param name="table">数据源</param>
        /// <returns>显示的表数据</returns>
        private DataTable GetShowTable(DataTable table)
        {
            if (table == null || table.Rows.Count < ShowSheetDataCount)
                return table;

            DataTable showTable = table.Clone();

            for (int i = 0; i < ShowSheetDataCount; i++)
            {
                showTable.ImportRow(table.Rows[i]);
            }

            return showTable;
        }

        /// <summary>
        /// 批量导入 回调方法
        /// </summary>
        /// <param name="result">异步操作返回结果</param>
        private void ImportCallBack(IAsyncResult result)
        {
            //获取所调用的委托对象
            Func<Dictionary<String, String>, Dictionary<String, Int32>> handler = (Func<Dictionary<String, String>, Dictionary<String, Int32>>)((AsyncResult)result).AsyncDelegate;
            Dictionary<String, Int32> res = handler.EndInvoke(result); //获取调用结果

            Invoke(new Action(() =>
            {
                btnImportBatchSheets.Enabled = true;   //跨线程改变控件状态
                pgbBatchImport.Value = 0;
                lblImportTableName.Text = @"  导入进度:";
            }));

            //所有表单导入明细
            StringBuilder allTableImportDetails = new StringBuilder();

            //导入失败的表单明细
            StringBuilder failedTableImportDetails = new StringBuilder();

            //导入失败的表单数量
            Int32 failedCount = 0;

            //是否已更新checktime
            Boolean hasUpdated = false;

            allTableImportDetails.AppendLine("本次导入明细如下 :");
            failedTableImportDetails.AppendLine("本次导入异常表格明细如下 :");

            //遍历每个表单导入结果,构造导入结果明细
            foreach (var item in res)
            {
                //添加到所有表单导入情况
                allTableImportDetails.AppendLine(String.Format("{0} 导入数量为: {1}", item.Key, item.Value));

                //导入结果为0的表单,添加到失败明细
                if (item.Value == 0)
                {
                    failedCount++;
                    failedTableImportDetails.AppendLine(String.Format("{0} 导入数量为: {1}", item.Key, item.Value));
                }
                else
                {
                    //更新template_checkinfo
                    if (!hasUpdated && (item.Key.StartsWith("b_") || item.Key.StartsWith("d_")))
                    {
                        ExcelBLL.UpdateCheckInfoTime();
                        hasUpdated = true;
                    }
                }
            }

            //重新加载数据
            if (ExcelBLL.IfCreateTable())
            {
                Invoke(new Action(() =>
                {
                    InitDbTables();
                }));
            }

            //记录所有表单导入明细
            Trace.Write(allTableImportDetails.ToString());

            if (failedTableImportDetails.Length > 16)
                Trace.Write(failedTableImportDetails.ToString());

            //如果全部导入成功 (包括 首页无数据的情况)
            if (failedCount == 0)
            {
                MessageBox.Show(@"导入成功!");
            }
            else
            {
                MessageBox.Show(failedTableImportDetails.ToString());
            }
        }

        /// <summary>
        /// 批量导入方法 (异步)
        /// </summary>
        /// <param name="selectedSheets">已选中的表单集合</param>
        /// <returns></returns>
        private Dictionary<String, Int32> Import(Dictionary<String, String> selectedSheets)
        {
            //保存导入结果信息
            Dictionary<String, Int32> importedInfos = new Dictionary<String, Int32>();

            //是否需要新建表
            Boolean ifCreateTable = ExcelBLL.IfCreateTable();

            //是否记录sql语句
            Boolean ifLogSQL = ConfigurationHelper.AppSettings["LogSql"].ToLower() == "true";

            //初始化进度条状态
            pgbBatchImport.Invoke(new Action(() =>
            {
                pgbBatchImport.Maximum = selectedSheets.Count;
            }));

            //遍历所有已选表单,进行导入
            foreach (String sheetName in selectedSheets.Keys)
            {
                pgbBatchImport.Invoke(new Action(() =>
                {
                    lblImportTableName.Text = sheetName;
                }));

                Int32 insertRowCount = 0;

                //构造sql执行语句
                List<String> sqlList = new List<String>();

                DataTable table = GetSheetTable(sheetName, selectedSheets[sheetName]);

                try
                {
                    if (table != null)
                    {
                        //构造sql执行语句
                        sqlList = ExcelBLL.GetSQL(table);

                        //如果表存在,则执行sql
                        if (AllDbTableNames.Contains(sheetName.ToLower()) || ifCreateTable)
                        {
                            insertRowCount = ExcelBLL.Insert(sqlList, sheetName, selectedSheets[sheetName], ifCreateTable);
                        }
                    }
                }
                catch (Exception ex)
                {
                    //导入失败时,设置导入结果
                    insertRowCount = 0;

                    //记录异常日志信息
                    StringBuilder sb = new StringBuilder();
                    sb.AppendLine(selectedSheets[sheetName]);
                    sb.AppendLine("异常表单: " + sheetName);
                    sb.AppendLine("生成的SQL: " + String.Concat(sqlList));
                    sb.AppendLine("异常信息: " + ex.Message);
                    sb.AppendLine("StackTrace: " + ex.StackTrace);
                    Trace.Write(sb.ToString());
                }
                finally
                {
                    //保存单个表单导入结果
                    importedInfos.Add(sheetName, insertRowCount);

                    //记录所有插入sql
                    if (ifLogSQL && table != null)
                    {
                        StringBuilder sb = new StringBuilder();
                        sb.AppendLine("#---------------------------------------" + sheetName + "-------------------------------------------");
                        sb.Append(String.Concat(sqlList));
                        sb.AppendLine();
                        sb.AppendLine();
                        Trace.Write(sb.ToString(), CurrentSqlFilePath);
                    }

                    //跨线程更新UI
                    pgbBatchImport.Invoke(new Action(() =>
                    {
                        pgbBatchImport.Value++;
                    }));
                }
            }

            return importedInfos;
        }

        /// <summary>
        /// 创建新TablePage标签
        /// </summary>
        /// <param name="dt">数据源</param>
        private void CreateNewTabPage(DataTable dt)
        {
            //删除已有的TabPage
            foreach (TabPage tabPage in tabControlSheetInfo.TabPages)
            {
                tabControlSheetInfo.TabPages.Remove(tabPage);
            }

            TabPage page = new TabPage(dt.TableName);

            DataGridView dgv = CreatSingleGridview(dt);

            page.Controls.Add(dgv);

            tabControlSheetInfo.TabPages.Add(page);
        }

        /// <summary>
        /// 主线程调用方法
        /// </summary>
        /// <param name="action">委托方法</param>
        private void InvokeMethod(Action action)
        {
            this.BeginInvoke(action);
        }

        #region Excel读取/缓存

        /// <summary>
        /// 加载Excel对象
        /// </summary>
        /// <param name="path">excel文档</param>
        /// <returns>Excel对象</returns>
        internal MoqikakaExcel LoadExcel(String path)
        {
            //优先读取缓存
            if (mExcels.ContainsKey(path))
            {
                MoqikakaExcel cache = mExcels[path];

                FileInfo info = new FileInfo(path);

                //文档没有修改过
                if (cache.ModifyDate == info.LastWriteTime)
                    return mExcels[path];

                //清理表表数据缓存
                foreach (var sheetName in cache.SheetNameList)
                {
                    if (mAllTables.Tables.Contains(sheetName))
                        mAllTables.Tables.Remove(sheetName);
                }
            }

            //重新加载
            MoqikakaExcel excel = new MoqikakaExcel(path);

            //缓存已读Excel文档对象 (并发插入异常?)
            lock (lockObj)
                mExcels[path] = excel;

            return excel;
        }

        /// <summary>
        /// 获取表单数据
        /// </summary>
        /// <param name="sheetName">表单名</param>
        /// <param name="filePath">文档路径</param>
        /// <returns>表单数据</returns>
        internal DataTable GetSheetTable(String sheetName, String filePath)
        {
            //获取excel文档对象
            MoqikakaExcel excel = LoadExcel(filePath);

            return GetSheetTable(sheetName, excel);
        }

        /// <summary>
        /// 获取表单数据
        /// </summary>
        /// <param name="sheetName">表单名</param>
        /// <param name="excel">文档对象</param>
        /// <returns>表单数据</returns>
        internal DataTable GetSheetTable(String sheetName, MoqikakaExcel excel)
        {
            //优先读取缓存数据
            if (mAllTables.Tables.Contains(sheetName))
                return mAllTables.Tables[sheetName];

            //没用的表单直接返回
            if (ExcelBLL.IsUselessSheet(sheetName))
                return null;

            //读取表单
            var table = ExcelBLL.TryRead(excel, sheetName);

            //加入缓存
            if (table != null)
                mAllTables.Tables.Add(table);

            return table;
        }

        #endregion

        #endregion
    }
}