// ****************************************
// FileName:ExcelDAL.cs
// Description:Excel业务,数据库支持类
// Tables:Many
// Author:Gavin && Burney
// Create Date:2014-06-01
// Revision History:
// ****************************************

using System;
using System.Data;

namespace DAL
{
    using MySql.Data.MySqlClient;
    using Utils.Configuration;

    public class ExcelDAL : BaseDAL
    {
        #region Excle导出

        /// <summary>
        /// 获取表所有数据
        /// </summary>
        /// <param name="tableName">表名</param>
        /// <returns>表数据</returns>
        public static DataTable GetTableData(String tableName)
        {
            String commandText = String.Format("SELECT * FROM {0}", tableName);
            return ExecuteDataTable(commandText);
        }

        /// <summary>
        /// 获取表字段和备注信息
        /// </summary>
        /// <param name="tableName">表名</param>
        /// <returns>表字段和备注信息</returns>
        public static DataTable GetComments(String tableName)
        {
            String dataBaseName = GetDataBaseNameFromConnectionString();
            String commandText = String.Format("SELECT column_name,column_comment,data_type FROM information_schema.columns WHERE table_schema='{0}' AND table_Name='{1}'", dataBaseName, tableName);
            return ExecuteDataTable(commandText);
        }

        /// <summary>
        /// 从数据库连接字符串中获取数据库名称
        /// </summary>
        /// <returns>数据库名称</returns>
        private static String GetDataBaseNameFromConnectionString()
        {
            var strArray = DbConnectionString.Split(';');
            String dataBaseName = null;
            foreach (String str in strArray)
            {
                if (str.ToLower().IndexOf("database") != -1)
                {
                    dataBaseName = str.Split('=')[1].Trim();
                }
            }
            return dataBaseName;
        }

        #endregion

        #region Excle导入

        /// <summary>
        /// 检测数据库是否可用
        /// </summary>
        /// <returns>如果数据库可连接返回true; 否则false</returns>
        public static Boolean IsDataBaseAccess()
        {
            try
            {
                ExecuteNonQuery("SELECT 1 FROM dual");
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// 执行无返回结果
        /// </summary>
        /// <param name="commandText">sql字符串</param>
        /// <returns>受影响行数</returns>
        public static Int32 ExecuteNonQuery(String commandText)
        {
            return BaseDAL.ExecuteNonQuery(commandText);
        }

        /// <summary>
        /// 获取数据库所有表名的集合
        /// </summary>
        /// <returns>表名的集合</returns>
        public static DataTable GetTableNames()
        {
            String dataBaseName = GetDataBaseNameFromConnectionString();
            String commandTxet = String.Format("SELECT table_name FROM information_schema.tables WHERE table_schema = '{0}';", dataBaseName);
            return ExecuteDataTable(commandTxet);
        }

        /// <summary>
        /// 清空表格数据
        /// </summary>
        /// <param name="tableName">表名</param>
        /// <returns>受影响行数</returns>
        public static Boolean TruncateTableData(String tableName)
        {
            String commandText = "TRUNCATE " + tableName;
            return ExecuteNonQuery(commandText) > 0;
        }

        /// <summary>
        /// 更新template_checkinfo表的时间
        /// </summary>
        public static void UpdateCheckInfoTime()
        {
            String commandText = "TRUNCATE template_checkinfo; Replace template_checkinfo SET UpdateTime = SYSDATE();";
            ExecuteNonQuery(commandText);
        }

        /// <summary>
        /// 获取template_checkinfo表的时间
        /// </summary>
        public static DateTime GetUpdateCheckInfoTime()
        {
            String commandText = "SELECT UpdateTime FROM template_checkinfo";
            return Convert.ToDateTime(ExecuteScalar(commandText));
        }

        #endregion
    }
}
