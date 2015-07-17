// ****************************************
// FileName:BaseOperation.cs
//// Description:Excel操作帮助类
// Tables:Nothing
// Author:Gavin
// Create Date:2014/11/25 18:10:00
// Revision History:
// ****************************************

using System;
using System.Data;
using System.IO;
using System.Collections.Generic;

namespace Utils.Excel
{
    using NPOI.SS.UserModel;
    using NPOI.XSSF.UserModel;
    using NPOI.HSSF.UserModel;

    /// <summary>
    /// Excel操作帮助类
    /// </summary>
    public static class ExcelHelper
    {
        #region 静态变量和初始化

        //Excel扩展名
        private const String ext2003 = ".xls";
        private const String ext2007 = ".xlsx";

        //NPOI最小日期
        private static DateTime minDate = new DateTime(1899, 12, 31);

        /// <summary>
        /// 数据参考列序号
        /// </summary>
        public static Int32 KeyColumnIndex { get; set; }

        /// <summary>
        /// 数据参考行序号
        /// </summary>
        public static Int32 KeyRowIndex { get; set; }

        /// <summary>
        /// 静态构造函数
        /// </summary>
        static ExcelHelper()
        {
            KeyColumnIndex = 0;
            KeyRowIndex = 0;
        }

        /// <summary>
        /// 设置读取表单时的参考序号
        /// </summary>
        /// <param name="rowIndex">参考行序号</param>
        /// <param name="columnIndex">参考列序号</param>
        public static void SetReferIndex(Int32 rowIndex, Int32 columnIndex)
        {
            KeyRowIndex = rowIndex;
            KeyColumnIndex = columnIndex;
        }

        #endregion

        #region 接口方法

        /// <summary>
        /// 加载Excel
        /// </summary>
        /// <param name="path">Excel文件路径</param>
        /// <returns>IWorkbook表单</returns>
        public static IWorkbook Load(String path)
        {
            if (!File.Exists(path))
                throw new Exception("找不到文件: " + path);

            String extension = Path.GetExtension(path);

            if (extension == ext2003)
                return new HSSFWorkbook(File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite));

            if (extension == ext2007)
                return new XSSFWorkbook(File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite));

            throw new Exception("文件格式错误!");
        }

        /// <summary>
        /// 读取表单数据
        /// </summary>
        /// <param name="workbook">Excel文档对象</param>
        /// <param name="sheetIndex">表单序号,从0开始</param>
        /// <returns>表单数据</returns>
        public static DataTable ReadAt(this IWorkbook workbook, Int32 sheetIndex)
        {
            if (sheetIndex > workbook.NumberOfSheets - 1)
                throw new Exception("表单序号越界!");

            ISheet sheet = workbook.GetSheetAt(sheetIndex);

            Int32 rowNum = ExactRowNum(sheet);      //实际行数
            Int32 colNum = ExactColumnNum(sheet);   //实际列数

            //获取初始化的Table
            DataTable table = InitTable(sheet.SheetName, rowNum, colNum);

            //遍历单元格为DataTable赋值
            for (int i = 0; i < rowNum; i++)
            {
                IRow row = sheet.GetRow(i);

                for (int j = 0; j < colNum; j++)
                {
                    table.Rows[i][j] = GetCellValue(row.GetCell(j));
                }
            }

            return table;
        }

        /// <summary>
        /// 获取所有表单名的集合
        /// </summary>
        /// <param name="workbook">excel对象</param>
        /// <returns>表单名的集合</returns>
        public static List<String> GetSheetNameList(this IWorkbook workbook)
        {
            List<String> sheetNames = new List<String>();

            for (int i = 0; i < workbook.NumberOfSheets; i++)
            {
                sheetNames.Add(workbook.GetSheetAt(i).SheetName);
            }

            return sheetNames;
        }

        #endregion

        #region 私有方法

        /// <summary>
        /// 初始化DataTable行
        /// </summary>
        /// <param name="tableName">表名</param>
        /// <param name="rowNum">行数</param>
        /// <param name="colNum">列数</param>
        /// <returns>初始化后的DataTable</returns>
        private static DataTable InitTable(String tableName, Int32 rowNum, Int32 colNum)
        {
            DataTable table = new DataTable(tableName);

            //初始化列
            for (int i = 0; i < colNum; i++)
            {
                table.Columns.Add();
            }

            //初始化行
            for (int i = 0; i < rowNum; i++)
            {
                table.Rows.Add(table.NewRow());
            }

            return table;
        }

        /// <summary>
        /// 实际列数
        /// </summary>
        /// <param name="sheet">表单</param>
        /// <returns>实际列数</returns>
        private static Int32 ExactColumnNum(this ISheet sheet)
        {
            //参考行
            IRow refRow = sheet.GetRow(KeyRowIndex);

            if (refRow == null) throw new Exception(sheet.SheetName + " 数据格式错误!");

            for (int i = refRow.Cells.Count - 1; i >= 0; i--)
            {
                if (refRow.Cells[i].CellType != CellType.Blank && GetCellValue(refRow.Cells[i]).ToString() != String.Empty)
                    return i + 1;
            }

            throw new Exception(sheet.SheetName + " 数据格式错误!"); ;
        }

        /// <summary>
        /// 实际行数
        /// </summary>
        /// <param name="sheet">表单</param>
        /// <returns>实际行数</returns>
        private static Int32 ExactRowNum(this ISheet sheet)
        {
            Int32 refRowNum = sheet.PhysicalNumberOfRows; //参考行数

            //从最后行开始读取,直到读取到关键参考列有值的行
            for (Int32 i = refRowNum; i >= 0; i--)
            {
                IRow row = sheet.GetRow(i);

                if (row == null || row.Cells.Count == 0)
                    continue;

                if (GetCellValue(row.Cells[KeyColumnIndex]).ToString() == String.Empty)
                    continue;

                return i + 1;
            }

            throw new Exception(sheet.SheetName + " 数据格式错误!");
        }

        /// <summary>
        /// 获取单元格数据
        /// </summary>
        /// <param name="cell">Excel单元格对象</param>
        /// <returns>单元格数据</returns>
        private static Object GetCellValue(ICell cell)
        {
            if (cell == null) return String.Empty;

            //如果单元格类型为公式 则取其实际类型
            CellType cellType = cell.CellType != CellType.Formula ? cell.CellType : cell.CachedFormulaResultType;

            switch (cellType)
            {
                case CellType.String:
                    return cell.StringCellValue;

                case CellType.Numeric:
                    if (IsTimeFormatted(cell)) return cell.DateCellValue.TimeOfDay;
                    return cell.NumericCellValue;

                case CellType.Boolean:
                    return cell.BooleanCellValue;

                default:
                    return String.Empty;
            }
        }

        /// <summary>
        /// 是否为时间格式
        /// </summary>
        /// <param name="cell">单元格对象</param>
        /// <returns>是否为时间格式</returns>
        private static Boolean IsTimeFormatted(ICell cell)
        {
            return DateUtil.IsCellDateFormatted(cell) && cell.DateCellValue.Date == minDate;
        }

        #endregion
    }
}