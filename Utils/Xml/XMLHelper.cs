// ****************************************
// FileName:XMLHelper
// Description: Excel表单字段中英文映射xml文件读写帮助类
// Tables:
// Author:Gavin
// Create Date:2014/6/6 14:11:29
// Revision History:
// ****************************************

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace Utils.Xml
{
    public class XMLHelper
    {
        #region 成员变量&&初始化

        //Xml文档
        private XElement _xml;

        //存放映射关系的文件(路径)
        private String _xmlFile;

        /// <summary>
        /// 初始化构造器
        /// </summary>
        /// <param name="xmlFile">xml文件路径</param>
        public XMLHelper(String xmlFile)
        {
            _xmlFile = xmlFile;
            _xml = XElement.Load(xmlFile);
        }

        #endregion

        #region 读取映射信息

        /// <summary>
        /// 获取Excel表单名映射的表名
        /// </summary>
        /// <param name="sheetName">Excel表单名</param>
        /// <returns>若存在,返回映射后的TableName;否则返回null</returns>
        public String GetTableNameMapping(String sheetName)
        {
            var tables = _xml.Element("Tables"); //获取Tables结点

            String mappingName = null;

            foreach (var table in tables.Elements())
            {
                if (table.Attribute("ExcelTableName").Value.ToUpper() == sheetName.ToUpper())
                {
                    mappingName = table.Attribute("DBTableName").Value;
                }
            }
            return mappingName;
        }

        /// <summary>
        /// 根据Excel表单名获取 字段映射关系
        /// </summary>
        /// <param name="sheetName">表单名</param>
        /// <returns>表单字段映射集合</returns>
        public Dictionary<String,String> GetTableMappingInfo(String sheetName)
        {
            Dictionary<String, String> mappings = new Dictionary<String, String>();

            var tables = _xml.Element("Tables");

            foreach (var table in tables.Elements())
            {
                if (table.Attribute("ExcelTableName").Value.ToUpper() == sheetName.ToUpper())
                {
                    foreach (var node in table.Element("Columns").Elements())
                    {
                        mappings.Add(node.Attribute("ExcelColumnName").Value, node.Attribute("DBColumnName").Value);
                    }
                }
            }
            return mappings;
        }

        /// <summary>
        /// 基本xml节点模板
        /// </summary>
        /// <returns>基本节点模板xml字符串</returns>
        public String GetTemplate()
        {
            XElement ele = new XElement("Table", new XElement("Columns"));
            return ele.ToString();
        }

        /// <summary>
        /// 根据表单名获取已存在映射关系的xml字符串
        /// </summary>
        /// <param name="sheetName">表单名</param>
        /// <returns>映射关系,xml元素字符串</returns>
        public String GetTableMappingXmlString(String sheetName)
        {
            var tables = _xml.Element("Tables");

            if (tables == null) return null;

            foreach (var table in tables.Elements())
            {
                if (table.Attribute("ExcelTableName").Value.ToUpper() == sheetName.ToUpper())  //如果存在,则返回xml字符串
                {
                    return table.ToString();
                }
            }

            return null;
        }

        /// <summary>
        /// 加载已映射表单列表
        /// </summary>
        /// <returns>已映射表单列表</returns>
        public List<String> LoadMappingedSheetNameList()
        {
            List<String> list = new List<String>();
            var tables = _xml.Element("Tables");  //table结点元素

            if (tables == null) return list;

            foreach (var table in tables.Elements())
            {
                list.Add(table.Attribute("ExcelTableName").Value.ToUpper());
            }

            return list;
        }

        #endregion

        #region 添加映射关系

        /// <summary>
        /// 添加表名映射关系
        /// </summary>
        /// <param name="sheetName">表单名</param>
        /// <param name="tableName">数据库表名</param>
        /// <param name="strXmlElement">映射关系,xml元素字符串</param>
        public String AddTableNameMapping(String sheetName,String tableName ,String strXmlElement)
        {
            XElement ele = XElement.Parse(strXmlElement);
            ele.SetAttributeValue("ExcelTableName", sheetName);
            ele.SetAttributeValue("DBTableName", tableName);
            return ele.ToString();
        }

        /// <summary>
        /// 添加字段映射
        /// </summary>
        /// <param name="strXmlElement">映射关系xml元素字符串</param>
        /// <param name="excelColumnName">Excel字段名</param>
        /// <param name="dbColumnName">数据库字段名</param>
        /// <returns>映射关系,xml元素字符串</returns>
        public String AddColumnMapping(String strXmlElement, String excelColumnName,String dbColumnName)
        {
            XElement ele = XElement.Parse(strXmlElement);
            XElement addElement = new XElement("Column");
            addElement.SetAttributeValue("ExcelColumnName",excelColumnName);
            addElement.SetAttributeValue("DBColumnName",dbColumnName);
            ele.Element("Columns").Add(addElement);
            return ele.ToString();
        }

        /// <summary>
        /// 清空字段映射关系
        /// </summary>
        /// <param name="strXmlElement">映射关系,xml元素字符串</param>
        /// <returns></returns>
        public String ClearColumnMappings(String strXmlElement)
        {
            XElement ele = XElement.Parse(strXmlElement);
            ele.Element("Columns").RemoveNodes();
            return ele.ToString();
        }

        /// <summary>
        /// 将映射关系添加到xml文件
        /// </summary>
        /// <param name="strXmlElement">映射关系,xml元素字符串</param>
        /// <returns></returns>
        public String AddMappingToXMLFile(String strXmlElement)
        {
            try
            {
                //字符串转换为xml
                XElement ele = XElement.Parse(strXmlElement);

                //保存到xml文件
                _xml.Element("Tables").Add(ele);
                _xml.Save(_xmlFile);

                return "保存成功!";
            }
            catch(Exception ex)
            {
                return "保存异常: " + ex.Message;
            }
        }

        /// <summary>
        /// 删除指定Excel表单映射关系
        /// </summary>
        /// <param name="sheetName">表单名称</param>
        public String DeleteTableMappingBySheetName(String sheetName)
        {
            var tables = _xml.Element("Tables");

            foreach (var table in tables.Elements())
            {
                if (table.Attribute("ExcelTableName").Value.ToUpper() == sheetName.ToUpper())  //如果存在,则返回xml字符串
                {
                    table.Remove();
                    //保存修改
                    _xml.Save(_xmlFile);

                    return "删除成功!";
                }
            }

            //保存修改
            _xml.Save(_xmlFile);

            return "未找到该映射信息!";
        }

        #endregion
    }
}
