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

        public static string ClickRightToCopy = "右击可复制表名";

        public static string ImportProccess = @"  导入进度:";

        public static string NoSelectedExcelToImport = "没有可导入的数据, 或配置中CreateTable == false!";

        public static string ConnectStringErrorTips = @"请检查该数据库连接字符串是否正确!";

        public static string ImportDetailResultTips = "本次导入明细如下 :";

        public static string ImorptErrorResultTips = "本次导入异常表格明细如下 :";

        public static string ImportResultInfo = "{0} 导入数量为: {1}";

        public static string ImportSuccess = @"导入成功!";
    }
}
