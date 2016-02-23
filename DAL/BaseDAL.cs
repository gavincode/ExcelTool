// ****************************************
// FileName:BaseDAL.cs
// Description:数据处理的基类
// Tables:Nothing
// Author:Jordan Zuo
// Create Date:2014-04-23
// Revision History:
// ****************************************

using System;
using System.Data;

namespace DAL
{
    using MySql.Data.MySqlClient;
    using Utils.Configuration;

    /// <summary>
    /// 数据处理的基类
    /// </summary>
    public class BaseDAL
    {
        //私有连接字符串
        private static String dbConnectionString;

        /// <summary>
        /// 静态构造函数
        /// </summary>
        static BaseDAL()
        {
            dbConnectionString = ConfigurationHelper.ConnectionString.Value;
        }

        /// <summary>
        /// 数据库连接字符串
        /// </summary>
        public static String DbConnectionString
        {
            get
            {
                return dbConnectionString;
            }
            set
            {
                dbConnectionString = value;
            }
        }

        /// <summary>
        /// 执行无返回值的数据库操作，仅返回受影响的行数
        /// </summary>
        /// <param name="commandText">sql语句</param>
        /// <param name="commandParameters">参数</param>
        /// <returns>受影响的行数</returns>
        protected static Int32 ExecuteNonQuery(String commandText, params MySqlParameter[] commandParameters)
        {
            try
            {
                return MySqlHelper.ExecuteNonQuery(DbConnectionString, commandText, commandParameters);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);  //Edit by gavin 2014-06-09  抛出异常
            }
        }

        /// <summary>
        /// 执行返回单个值的数据库操作
        /// </summary>
        /// <param name="commandText">sql语句</param>
        /// <param name="commandParameters">参数</param>
        /// <returns>返回单个值</returns>
        protected static Object ExecuteScalar(String commandText, params MySqlParameter[] commandParameters)
        {
            try
            {
                return MySqlHelper.ExecuteScalar(DbConnectionString, commandText, commandParameters);
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        /// <summary>
        /// 执行数据库操作，返回一个数据表
        /// </summary>
        /// <param name="commandText">sql语句</param>
        /// <param name="commandParameters">参数</param>
        /// <returns>包含结果的数据表</returns>
        protected static DataTable ExecuteDataTable(String commandText, params MySqlParameter[] commandParameters)
        {
            try
            {
                DataSet ds = MySqlHelper.ExecuteDataset(DbConnectionString, commandText, commandParameters);
                if (ds != null && ds.Tables.Count > 0)
                {
                    return ds.Tables[0];
                }
            }
            catch (Exception ex)
            {
                return null;
            }

            return null;
        }
    }
}