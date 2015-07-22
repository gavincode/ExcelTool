// ****************************************
// FileName:XMLHelper
// Description: Excel表单和数据库表格映射页面
// Tables:Many
// Author:Gavin
// Create Date:2014/6/6 14:11:29
// Revision History:
// ****************************************

using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelTool
{
    using Utils.Excel;
    using Utils.Xml;
    using BLL;
    public partial class TableMapping : Form
    {
        #region 01 成员变量&&初始化

        //父窗体
        private Main _parent;

        //当前页面所需映射的表单名
        private String _sheetName;

        //当前操作的Excel文件路径
        private String _excelFile;

        //Xml读取和写入帮助类
        private XMLHelper _xmlHelper = new XMLHelper("TableMapping.xml");

        /// <summary>
        /// 初始化构造器
        /// </summary>
        /// <param name="sheetName">表单名</param>
        /// <param name="excelFile">Excel文件路径</param>
        /// <param name="parent">父窗体</param>
        public TableMapping(String sheetName, String excelFile, Main parent)
        {
            InitializeComponent();

            //初始化成员变量
            _sheetName = sheetName;
            _excelFile = excelFile;
            txtSheeName.Text = sheetName;
            _parent = parent;

            //初始化数据库表数据
            InitDbTables();

            //初始化页面控件
            InitMappingText(sheetName, true);

            //绑定Excel表单的字段信息

            //设置ListView样式
            InitListViewStyle();
        }

        #endregion

        #region 02 事件

        /// <summary>
        /// 添加字段映射关系
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnADD_Click(object sender, EventArgs e)
        {
            //ListView中没有CheckedItems时,不允许添加
            if (listViewExcelColumn.CheckedItems.Count < 1) return;
            if (listViewDbColumn.CheckedItems.Count < 1) return;

            //获取所选映射关系的字段序号
            Int32 excelColumnIndex = listViewExcelColumn.CheckedItems[0].Index;
            Int32 dbColumnIndex = listViewDbColumn.CheckedItems[0].Index;

            //获取所选映射关系的字段名
            String excelColumnName = listViewExcelColumn.CheckedItems[0].Text;
            String dbColumnName = listViewDbColumn.CheckedItems[0].Text;

            //移除已添加项
            listViewDbColumn.CheckedItems[0].Remove();
            listViewExcelColumn.CheckedItems[0].Remove();

            //设置默认Check项
            if (listViewDbColumn.Items.Count > 0)
            {
                if (dbColumnIndex >= listViewDbColumn.Items.Count)
                {
                    listViewDbColumn.Items[listViewDbColumn.Items.Count - 1].Checked = true;
                }
                else
                {
                    listViewDbColumn.Items[dbColumnIndex].Checked = true;
                }
            }
            if (listViewExcelColumn.Items.Count > 0)
            {
                if (excelColumnIndex >= listViewExcelColumn.Items.Count)
                {
                    listViewExcelColumn.Items[listViewExcelColumn.Items.Count - 1].Checked = true;
                }
                else
                {
                    listViewExcelColumn.Items[excelColumnIndex].Checked = true;
                }
            }

            //显示添加结果
            rtxtMapping.Text = _xmlHelper.AddColumnMapping(rtxtMapping.Text, excelColumnName, dbColumnName);

            //设置保存按钮状态
            if (listViewExcelColumn.Items.Count == 0) btnSave.Enabled = true;

            //添加后才能重置
            btnReset.Enabled = true;
        }

        /// <summary>
        /// 数据库表选择改变事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cmbDbTables_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbDbTables.SelectedIndex < 0) return;

            var list = ExcelBLL.GetComments(cmbDbTables.Text);

            InitDbColumn(list);  //初始化Excel字段名列表

            InitExcelColumn();   //初始化所选数据库表字段名列表

            //显示映射关系
            InitMappingText(txtSheeName.Text, false);
        }

        /// <summary>
        /// 数据库字段ListView单选
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listViewDbColumn_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            //只能单选
            if (e.Item.Checked)
            {
                foreach (ListViewItem item in listViewDbColumn.CheckedItems)
                {
                    if (item.Checked && item != e.Item)
                        item.Checked = false;
                }
            }
        }

        /// <summary>
        /// Excel表单字段ListView单选
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listViewExcelColumn_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            //只能单选
            if (e.Item.Checked)
            {
                foreach (ListViewItem item in listViewExcelColumn.CheckedItems)
                {
                    if (item.Checked && item != e.Item) item.Checked = false;
                }
            }
        }

        /// <summary>
        /// 保存映射关系
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSave_Click(object sender, EventArgs e)
        {
            //防止多次添加
            btnSave.Enabled = false;

            //保存映射关系
            String res = _xmlHelper.AddMappingToXMLFile(rtxtMapping.Text.Trim());

            //显示并关闭弹出窗口
            MessageBox.Show(res);
            DialogResult = DialogResult.OK;
            Close();
        }

        /// <summary>
        /// 重置映射关系
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnReset_Click(object sender, EventArgs e)
        {
            cmbDbTables_SelectedIndexChanged(null, null);
        }

        /// <summary>
        /// 删除已有映射
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnDelete_Click(object sender, EventArgs e)
        {
            //删除文档中的映射关系
            String res = _xmlHelper.DeleteTableMappingBySheetName(txtSheeName.Text);

            //初始化页面映射信息
            InitMappingText(txtSheeName.Text, false);

            MessageBox.Show(res);
        }

        #endregion

        #region 03 辅助方法

        ///  <summary>
        ///  初始化页面映射关系相关信息
        ///  </summary>
        /// <param name="sheetName">表单名</param>
        /// <param name="loadFromXml">若要从xml映射文件中读取映射关系,true;否则 false</param>
        private void InitMappingText(String sheetName, Boolean loadFromXml)
        {
            //不从映射文件读取
            if (!loadFromXml)
            {
                String temp = _xmlHelper.GetTemplate();
                rtxtMapping.Text = _xmlHelper.AddTableNameMapping(sheetName, cmbDbTables.Text, temp);
            }
            //初始化时,绑定已存在映射关系(如果有)
            else
            {
                //读取文档中的映射信息
                String xmlMappingString = _xmlHelper.GetTableMappingXmlString(sheetName);

                if (xmlMappingString == null)
                {
                    String temp = _xmlHelper.GetTemplate();
                    rtxtMapping.Text = _xmlHelper.AddTableNameMapping(sheetName, cmbDbTables.Text, temp);
                }
                else
                {
                    //设置已存在映射关系的数据库表
                    String mappingTableName = _xmlHelper.GetTableNameMapping(sheetName);
                    cmbDbTables.Text = mappingTableName;
                    rtxtMapping.Text = xmlMappingString;

                    //存在映射关系时,仅删除按钮可用
                    cmbDbTables.Enabled = false;
                    btnADD.Enabled = false;
                    btnReset.Visible = false;
                    btnDelete.Visible = true;
                    return;
                }
            }

            //不存在映射关系时, 恢复按钮状态
            cmbDbTables.Enabled = true;
            btnADD.Enabled = true;
            btnReset.Visible = true;
            btnDelete.Visible = false;
        }

        /// <summary>
        /// 初始化Excel表单的字段信息列表
        /// </summary>
        private void InitExcelColumn()
        {
            //清空列表
            listViewExcelColumn.Clear();

            Form form = this.Owner;

            //从Excel中读取该表单的字段列表
            List<String> list = CommonBLL.GetCloumnList(GlobalCacheBLL.GetSheetTable(_sheetName, _excelFile));

            if (list == null || list.Count == 0) return;

            //根据表单字段名,添加listViewExcelColumn节点
            foreach (String item in list)
            {
                ListViewItem listItem = new ListViewItem(item) { Checked = false };
                listViewExcelColumn.Items.Add(listItem);
            }

            //设置默认Check项
            listViewExcelColumn.Items[0].Checked = true;
        }

        /// <summary>
        /// 初始化数据库表的字段信息
        /// </summary>
        /// <param name="list"></param>
        private void InitDbColumn(Dictionary<String, String[]> list)
        {
            //清空列表
            listViewDbColumn.Clear();

            if (list.Keys.Count == 0) return;

            //根据数据库字段,添加listViewDbColumn节点
            foreach (var item in list.Keys)
            {
                ListViewItem listItem = new ListViewItem(item) { Checked = false };
                listViewDbColumn.Items.Add(listItem);
            }

            //设置默认Check项
            listViewDbColumn.Items[0].Checked = true;
        }

        /// <summary>
        /// 初始化ListView的样式
        /// </summary>
        private void InitListViewStyle()
        {
            //设置ListView的高度
            ImageList imgList = new ImageList { ImageSize = new Size(1, 16) };
            listViewExcelColumn.SmallImageList = imgList;
            listViewDbColumn.SmallImageList = imgList;
        }

        /// <summary>
        /// 初始化数据库数据表名List
        /// </summary>
        private void InitDbTables()
        {
            //数据库是否可用
            if (!ExcelBLL.IsDataBaseAccess()) return;

            cmbDbTables.DataSource = ExcelBLL.GetTableNameList();
        }

        #endregion
    }
}