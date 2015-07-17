// ****************************************
// FileName:IgnoreSheetsBLL.cs
// Description:忽略表单处理类
// Tables:
// Author:Gavin
// Create Date:2014/12/19 11:15:05
// Revision History:
// ****************************************

using System;
using System.Data;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using System.IO;

namespace BLL
{
    /// <summary>
    /// 忽略表单处理类
    /// </summary>
    public class IgnoreSheetsBLL
    {
        //忽略表单列表
        private static readonly List<String> IgnoreSheetList = new List<String>();

        //配置文件
        private static readonly String mIgnoreFilePath = AppDomain.CurrentDomain.BaseDirectory + "\\IgnoreSheets.txt";

        /// <summary>
        /// 重置忽略表单列表
        /// </summary>
        public static void Reset()
        {
            if (!File.Exists(mIgnoreFilePath))
                File.CreateText(mIgnoreFilePath).Close();

            //清空列表
            IgnoreSheetList.Clear();

            //读取列表
            String[] sheets = File.ReadAllLines(mIgnoreFilePath, Encoding.Default);

            foreach (var sheetName in sheets)
            {
                IgnoreSheetList.Add(sheetName.Trim().ToLower());
            }
        }

        /// <summary>
        /// 是否为忽略的表单
        /// </summary>
        /// <param name="sheetName">表单名</param>
        /// <returns>是否为忽略的表单</returns>
        public static Boolean IsIgnoreSheet(String sheetName)
        {
            return IgnoreSheetList.Contains(sheetName.ToLower());
        }
    }
}
