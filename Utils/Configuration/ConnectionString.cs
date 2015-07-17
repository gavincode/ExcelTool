// ****************************************
// FileName:ConnectionString.cs
// Description:ConnectionString读取帮助类
// Tables:None
// Author:Gavin && Burney
// Create Date:2014-06-01
// Revision History:
// ****************************************

using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace Utils.Configuration
{
    /// <summary>
    /// ConnectionString数据库连接字符串读取帮助类
    /// </summary>
    public class ConnectionString
    {
        //默认Config文件名
        private String _defaultConfig = String.Empty;
        private String _defaultConnectionName = "DBConnectionString";

        public ConnectionString()
        {
        }

        /// <summary>
        /// 构造方法
        /// </summary>
        /// <param name="configFile">Config文件名</param>
        public ConnectionString(String configFile)
        {
            this._defaultConfig = configFile;
        }

        /// <summary>
        /// 获取或设置ConnectionString 配置信息
        /// </summary>
        public String Value
        {
            get
            {
                //以Xml类型加载配置文件
                XmlDocument config = new XmlDocument();
                config.Load(_defaultConfig);

                XmlNode connectionStringNode = config.SelectSingleNode("configuration/connectionStrings");
                if (connectionStringNode == null) return String.Empty;

                XmlNodeList addNodeList = connectionStringNode.SelectNodes("add");

                //遍历获取连接字符串的值
                foreach (XmlNode xmlNode in addNodeList)
                {
                    if (xmlNode.Attributes["name"].Value == _defaultConnectionName)
                    {
                        return xmlNode.Attributes["connectionString"].Value;
                    }
                }

                return String.Empty;
            }
            set
            {
                //以Xml类型加载配置文件
                XmlDocument config = new XmlDocument();
                config.Load(_defaultConfig);

                XmlNode connectionStringNode = config.SelectSingleNode("configuration/connectionStrings");
                XmlNodeList addNodeList = connectionStringNode.SelectNodes("add");

                //修改连接字符串的值
                foreach (XmlNode xmlNode in addNodeList)
                {
                    if (xmlNode.Attributes["name"].Value == _defaultConnectionName)
                    {
                        xmlNode.Attributes["connectionString"].Value = value;
                        break;
                    }
                }

                config.Save(this._defaultConfig);
            }
        }
    }
}
