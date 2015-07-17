// ****************************************
// FileName:MySqlConnectingBase.cs
// Description:提供MySql连接字符串
// Tables:None
// Author:Gavin && Burney
// Create Date:2014-06-01
// Revision History:
// ****************************************

using System;
using System.Collections.Generic;
using System.Configuration;
using System.Text;

namespace ExcelTool.DAL
{
    //using Moqikaka.Util;

    /// <summary>
    /// 提供MySql连接字符串
    /// </summary>
    public static class MySqlConnectingBase
    {
        /// <summary>
        /// MySql数据库连接字符串
        /// </summary>
        public static  String ConnectString
        {
            get
            {
                return Moqikaka.Util.AppConfigUtil.GetConnectionStringsConfig("MySqlStr");
            }
        }
    }
}
