// ****************************************
// FileName:ConfigurationHelper.cs
// Description:Configuration读取操作帮助类
// Tables:None
// Author:Gavin && Burney
// Create Date:2014-06-01
// Revision History:
// ****************************************

using System;
using System.Collections.Generic;
using System.Text;

namespace Utils.Configuration
{
    /// <summary>
    /// Configuration读取操作帮助类
    /// </summary>
    public static class ConfigurationHelper
    {
        //默认config文件
        private static String _defaultConfig = "ExcelTool.exe.config";

        /// <summary>
        /// 连接字符串帮助类
        /// </summary>
        public static ConnectionString ConnectionString = null;

        /// <summary>
        /// AppSettings帮助类
        /// </summary>
        public static AppSettings AppSettings = null;

        /// <summary>
        /// 静态属性获取或设置Config的文件名
        /// </summary>
        public static String ConfigFile
        {
            get
            {
                return _defaultConfig;
            }

            set
            {
                _defaultConfig = value;
            }
        }

        /// <summary>
        /// 静态构造方法
        /// </summary>
        static ConfigurationHelper()
        {
            Init();
        }

        /// <summary>
        /// 初始化
        /// </summary>
        public static void Init()
        {
            ConnectionString = new ConnectionString(_defaultConfig);
            AppSettings = new AppSettings(_defaultConfig);
        }
    }
}
