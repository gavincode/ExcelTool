using System;
using System.Collections.Generic;
using System.Web;
using System.Data;

namespace Utils.Excel
{
    /// <summary>
    /// 将DataTable转换为插入sql语句
    /// </summary>
    public interface IDataTableToSql
    {
        /// <summary>
        /// DataTableToSql
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        string DataTableToSql(DataTable dt);
    }
}