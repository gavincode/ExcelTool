// ****************************************
// FileName:CommonBLL.cs
// Description:一般逻辑处理类
// Tables:Nothing
// Author:Gavin
// Create Date:2014/12/2 10:09:52
// Revision History:
// ****************************************

using System;
using System.Data;
using System.Linq;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace BLL
{
    using Utils.Configuration;

    /// <summary>
    /// 一般逻辑处理类
    /// </summary>
    public class CommonBLL
    {
        /// <summary>
        /// 当前应用程序主目录
        /// </summary>
        private static String AppDomin
        {
            get
            {
                return System.AppDomain.CurrentDomain.BaseDirectory;
            }
        }

        /// <summary>
        /// 获取当前SQL文件路径
        /// </summary>
        /// <returns>当前sql文件路径</returns>
        public static String GetCurrentSqlFilePath()
        {
            String sqlFolder = AppDomin + "\\SQL";

            Int32 maxIndexOfToday = Directory.GetFiles(sqlFolder, DateTime.Now.ToString("yyyy-MM-dd") + "*").Length;

            String nextIndex = (maxIndexOfToday + 1).ToString().PadLeft(3, '0');

            return String.Format("{0}\\SQL\\{1}-{2}.txt", AppDomin, DateTime.Now.ToString("yyyy-MM-dd"), nextIndex);
        }

        /// <summary>
        /// 获取当前Log文件路径
        /// </summary>
        /// <returns>当前sql文件路径</returns>
        public static String GetCurrentLogFilePath()
        {
            return String.Format("{0}\\Log\\log{1}.txt", AppDomin, DateTime.Now.ToString("yyyy-MM-dd"));
        }

        /// <summary>
        /// 是否匹配中文
        /// </summary>
        /// <param name="text">文本</param>
        /// <returns>是否匹配中文</returns>
        public static Boolean MatchedChinese(String text)
        {
            return Regex.IsMatch(text, @"[\u4e00-\u9fa5]");
        }

        /// <summary>
        /// 获取对话框标题
        /// </summary>
        /// <returns>对话框标题</returns>
        public static String GetDialogTitle()
        {
            //上次导入时间
            var lastUpdateTime = ExcelBLL.GetUpdateCheckInfoTime();

            if (lastUpdateTime == DateTime.MinValue)
                return "打开";

            return String.Format("上一次导入时间: {0}", lastUpdateTime.ToString("yyyy-MM-dd HH:mm:ss"));
        }

        /// <summary>
        /// 获取上一次导出Excel文件夹路径
        /// </summary>
        /// <returns>上一次导出Excel文件夹路径</returns>
        public static String GetStoredFolder()
        {
            return ConfigurationHelper.AppSettings["ExportStoredFolder"];
        }

        /// <summary>
        /// 保存本次存放Excel的文件夹路径
        /// </summary>
        /// <param name="exportExcelFolder">导出Excel文件路径</param>
        public static void StoreExportFolder(String exportExcelFolder)
        {
            ConfigurationHelper.AppSettings["ExportStoredFolder"] = exportExcelFolder;
        }

        /// <summary>
        /// 获取表字段列表
        /// </summary>
        /// <param name="table">表</param>
        /// <returns>表字段列表</returns>
        public static List<String> GetCloumnList(DataTable table)
        {
            List<String> list = new List<String>();

            foreach (DataColumn item in table.Columns)
            {
                list.Add(item.ColumnName);
            }

            return list;
        }
    }
}
