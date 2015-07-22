//***********************************************************************************
// 文件名称：GlobalCacheBLL.cs
// 功能描述：全局缓存处理类
// 数据表：
// 作者：Gavin
// 日期：2015/7/22 15:50:58
// 修改记录：
//***********************************************************************************

using System;
using System.Data;
using System.Linq;
using System.Collections.Generic;

namespace BLL
{
    using Utils.Excel;
    using System.IO;

    /// <summary>
    /// 全局缓存处理类
    /// </summary>
    public class GlobalCacheBLL
    {
        private static readonly Object lockObj = new Object();

        //缓存-已读取过的Excel对象
        public static Dictionary<String, MoqikakaExcel> mExcels = new Dictionary<String, MoqikakaExcel>();

        //缓存-已读取的Excel表单
        public static DataSet mAllTables = new DataSet();

        /// <summary>
        /// 加载Excel对象
        /// </summary>
        /// <param name="path">excel文档</param>
        /// <returns>Excel对象</returns>
        public static MoqikakaExcel LoadExcel(String path)
        {
            //优先读取缓存
            if (mExcels.ContainsKey(path))
            {
                MoqikakaExcel cache = mExcels[path];

                //文档没有修改过
                if (cache.ModifyDate == File.GetLastWriteTime(path))
                    return mExcels[path];

                //清理表表数据缓存
                foreach (var sheetName in cache.SheetNameList)
                {
                    if (mAllTables.Tables.Contains(sheetName))
                        mAllTables.Tables.Remove(sheetName);
                }
            }

            //重新加载
            MoqikakaExcel excel = new MoqikakaExcel(path);

            //缓存已读Excel文档对象 (并发插入异常?)
            lock (lockObj)
                mExcels[path] = excel;

            return excel;
        }

        /// <summary>
        /// 获取表单数据
        /// </summary>
        /// <param name="sheetName">表单名</param>
        /// <param name="filePath">文档路径</param>
        /// <returns>表单数据</returns>
        public static DataTable GetSheetTable(String sheetName, String filePath)
        {
            //获取excel文档对象
            MoqikakaExcel excel = LoadExcel(filePath);

            return GetSheetTable(sheetName, excel);
        }

        /// <summary>
        /// 获取表单数据
        /// </summary>
        /// <param name="sheetName">表单名</param>
        /// <param name="excel">文档对象</param>
        /// <returns>表单数据</returns>
        public static DataTable GetSheetTable(String sheetName, MoqikakaExcel excel)
        {
            //优先读取缓存数据
            if (mAllTables.Tables.Contains(sheetName))
                return mAllTables.Tables[sheetName];

            //没用的表单直接返回
            if (ExcelBLL.IsUselessSheet(sheetName))
                return null;

            //读取表单
            var table = ExcelBLL.TryRead(excel, sheetName);

            //加入缓存
            if (table != null)
                mAllTables.Tables.Add(table);

            return table;
        }
    }
}
