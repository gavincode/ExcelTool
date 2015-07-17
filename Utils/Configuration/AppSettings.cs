// ****************************************
// FileName:AppSettings.cs
// Description:AppSetting读取帮助类
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
    /// AppSetting读取帮助类
    /// </summary>
    public class AppSettings
    {
        //默认配置文件
        private String _defaultConfig = String.Empty;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="configFile">配置文件路径</param>
        public AppSettings(String configFile)
        {
            this._defaultConfig = configFile;
        }

        /// <summary>
        /// 获取所有设置的键值对
        /// </summary>
        public Dictionary<String, String> AllSettings
        {
            get
            {
                //以Xml类型加载配置文件
                XmlDocument config = new XmlDocument();
                config.Load(_defaultConfig);

                XmlNode appSettingsNode = config.SelectSingleNode("configuration/appSettings");
                XmlNodeList addNodeList = appSettingsNode.SelectNodes("add");

                Dictionary<String, String> dict = new Dictionary<String, String>();

                //遍历所有的add节点
                foreach (XmlNode node in addNodeList)
                {
                    if (!dict.ContainsKey(node.Attributes["key"].Value))
                    {
                        dict[node.Attributes["key"].Value] = node.Attributes["value"].Value;
                    }
                }

                return dict;
            }
        }

        /// <summary>
        /// 索引器(根据keyName获取或设置keyValue)
        /// </summary>
        /// <param name="keyName"></param>
        /// <returns></returns>
        public String this[String keyName]
        {
            get
            {
                //以Xml类型加载配置文件
                XmlDocument config = new XmlDocument();
                config.Load(_defaultConfig);

                XmlNode appSettingsNode = config.SelectSingleNode("configuration/appSettings");
                XmlNodeList addNodeList = appSettingsNode.SelectNodes("add");

                //遍历所有的add节点
                foreach (XmlNode node in addNodeList)
                {
                    if (node.Attributes["key"] != null && node.Attributes["key"].Value == keyName)
                    {
                        return node.Attributes["value"].Value;
                    }
                }

                return null;
            }

            set
            {
                //以Xml类型加载配置文件
                XmlDocument config = new XmlDocument();
                config.Load(_defaultConfig);

                XmlNode appSettingsNode = config.SelectSingleNode("configuration/appSettings");
                XmlNodeList addNodeList = appSettingsNode.SelectNodes("add");

                //遍历所有的add节点
                foreach (XmlNode node in addNodeList)
                {
                    if (node.Attributes["key"] != null && node.Attributes["key"].Value == keyName)
                    {
                        node.Attributes["value"].Value = value;
                        break;
                    }
                }

                config.Save(this._defaultConfig);
            }
        }
    }
}
