// ****************************************
// FileName:ExcelOperationBase.cs
// Description:Excel基类
// Tables:None
// Author:Gavin && Burney
// Create Date:2014-06-01
// Revision History:
// ****************************************

using System;
using System.Collections.Generic;
using System.Data;

namespace Utils.Excel
{
    using NPOI.SS.UserModel;

    /// <summary>
    /// Excel基类
    /// </summary>
    public abstract class ExcelBase
    {
        #region 属性

        #region Protected

        /// <summary>
        /// Excel文档对象 (NPOI)
        /// </summary>
        protected IWorkbook Workbook { get; set; }

        /// <summary>
        /// Excel文件路径
        /// </summary>
        public String Path { get; private set; }

        #endregion

        #region Public

        /// <summary>
        /// 表单数量
        /// </summary>
        public Int32 NumberOfSheets
        {
            get
            {
                return Workbook.NumberOfSheets;
            }
        }

        /// <summary>
        /// 获取所有表单名的集合
        /// </summary>
        /// <returns>表单名的集合</returns>
        public List<String> SheetNameList
        {
            get
            {
                return Workbook.GetSheetNameList();
            }
        }

        #endregion

        #endregion

        #region 初始化

        /// <summary>
        /// Excel文件路径
        /// </summary>
        protected ExcelBase(String path)
        {
            this.Path = path;

            this.Workbook = ExcelHelper.Load(path);
        }

        #endregion

        #region 公开方法

        #region 读取

        /// <summary>
        /// 读取Excel表单
        /// </summary>
        /// <param name="sheetIndex">表单Index</param>
        /// <returns>表单数据</returns>
        public virtual DataTable ReadAt(Int32 sheetIndex)
        {
            return Workbook.ReadAt(sheetIndex);
        }

        /// <summary>
        /// 读取Excel表单
        /// </summary>
        /// <param name="sheetIndex">表单Index</param>
        /// <returns>表单数据或者null</returns>
        public virtual DataTable TryReadAt(Int32 sheetIndex)
        {
            try
            {
                return Workbook.ReadAt(sheetIndex);
            }
            catch
            {
                return null;
            }
        }

        #endregion

        #region 常用

        /// <summary>
        /// 获取表单名称
        /// </summary>
        /// <param name="sheetIndex">表单序号</param>
        /// <returns>表单名称</returns>
        public String GetSheetName(Int32 sheetIndex)
        {
            return Workbook.GetSheetName(sheetIndex);
        }

        /// <summary>
        /// 获取表单Index
        /// </summary>
        /// <param name="sheetName">表单名</param>
        /// <returns>表单序号</returns>
        public Int32 GetSheetIndex(String sheetName)
        {
            return Workbook.GetSheetIndex(sheetName);
        }

        #endregion

        #endregion
    }
}