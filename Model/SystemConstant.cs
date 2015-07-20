//***********************************************************************************
// 文件名称：SystemConstant.cs
// 功能描述：
// 数据表：
// 作者：Gavin
// 日期：2015/7/20 10:06:08
// 修改记录：
//***********************************************************************************

using System;
using System.Data;
using System.Linq;
using System.Collections.Generic;

namespace Model
{
    public class ConstantText
    {
        public static string ClickToChoose = "点击选择需要导入的Excel文档";

        public static string NoMappingRelation = "该表存在映射关系";

        public static string TableNotExist = "当前数据库不存在该表";

        public static string ClickToCopy = "点击复制表名";

        public static string ExcelFileFilter = @"Excel文档|*.xls;*.xlsx";

        public static string Import = "导入";

        public static string TestPlease = "请测试";

        public static string Find = @"查找";

        public static string AddComment = "添加表备注";

        public static string AddCommentTips = "请输入创建表时的备注名称,\r\n以Enter键保存(C启动)";

        public static string CopySheetName = "复制表名";

        public static string MappingTable = "映射该表";

        public static string MappingExist = "该表存在映射关系";

        public static string ExportDescription = @"Excel文档存放位置";

        public static string ShowFolder = @"查看文件夹";

        public static string ExportSuccessTips = "成功导出{0}个表格";

        public static string ShowAllSheets = "查看整个Excel表单";
    }
}
