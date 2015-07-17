// ****************************************
// FileName:MoqikakaExcel.cs
// Description:摩奇卡卡Excel文档操作类
// Tables:None
// Author:Gavin
// Create Date:2014-06-01
// Revision History:
// ****************************************

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace Utils.Excel
{
    using NPOI.SS.UserModel;
    using NPOI.XSSF.UserModel;

    /// <summary>
    /// 摩奇卡卡Excel文档操作类
    /// </summary>
    public class MoqikakaExcel : ExcelBase
    {
        #region 属性

        /// <summary>
        /// Excel文档更改日志
        /// </summary>
        public DateTime ModifyDate { get; set; }

        #endregion

        #region 初始化

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="path">Excel文档路径</param>
        public MoqikakaExcel(String path)
            : base(path)
        {
            FileInfo fileInfo = new FileInfo(path);
            ModifyDate = fileInfo.LastWriteTime;
        }

        /// <summary>
        /// 静态构造函数
        /// </summary>
        static MoqikakaExcel()
        {
            //设置读取表单的参考行/列序号
            ExcelHelper.SetReferIndex(1, 0);
        }

        #endregion

        #region 读取

        /// <summary>
        /// 读取表单数据
        /// </summary>
        /// <param name="sheetIndex">表单Index</param>
        /// <returns>表单数据</returns>
        public override DataTable ReadAt(Int32 sheetIndex)
        {
            //正常读取表单数据
            DataTable table = base.TryReadAt(sheetIndex);

            return HandleTable(table);
        }

        /// <summary>
        /// 处理表单数据格式
        /// </summary>
        /// <param name="table">数据源</param>
        /// <returns>处理后的数据</returns>
        private DataTable HandleTable(DataTable table)
        {
            if (table == null) return null;

            //表单字段描述行序号
            Int32 descRowNum = MoqikakaExcelSettings.DescRowNum;

            //构造表单列名
            if (descRowNum != -1)
            {
                DataRow descRow = table.Rows[descRowNum];
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    table.Columns[i].ColumnName = descRow[i].ToString();
                }
            }

            table.Rows.RemoveAt(descRowNum);

            return table;
        }

        #endregion

        #region 写入

        /// <summary>
        /// 写入Excel
        /// </summary>
        /// <param name="dataSource">数据源</param>
        /// <param name="path">Excel文件路径</param>
        /// <param name="paras">可选参数</param>
        public static void Write(DataTable dataSource, String path, params object[] paras)
        {
            //新建Excel表单
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet(dataSource.TableName);

            //前三行表单Header样式
            ICellStyle headerCellStyle = workbook.CreateCellStyle();
            headerCellStyle.VerticalAlignment = VerticalAlignment.Center;
            headerCellStyle.Alignment = HorizontalAlignment.Center;
            IFont font2 = workbook.CreateFont();
            font2.Boldweight = (short)FontBoldWeight.Normal;
            font2.FontName = "Consolas";
            font2.FontHeightInPoints = 11;
            font2.Color = (short)FontColor.Red;
            headerCellStyle.SetFont(font2);

            //构造表单第一行 备注信息
            Dictionary<String, String[]> commentDictionary = paras[0] as Dictionary<String, String[]>;
            IRow row0 = sheet.CreateRow(0);
            for (Int32 i = 0; i < dataSource.Columns.Count; i++)
            {
                ICell cell = row0.CreateCell(i);
                cell.CellStyle = headerCellStyle;
                cell.SetCellValue(commentDictionary[dataSource.Columns[i].Caption][0]);
            }

            //构造表单第二行 字段名
            IRow row1 = sheet.CreateRow(1);
            for (Int32 i = 0; i < dataSource.Columns.Count; i++)
            {
                ICell cell = row1.CreateCell(i);
                cell.CellStyle = headerCellStyle;
                cell.SetCellValue(dataSource.Columns[i].Caption);
            }

            //构造表单第三行 字段类型
            IRow row2 = sheet.CreateRow(2);
            for (Int32 i = 0; i < dataSource.Columns.Count; i++)
            {
                ICell cell = row2.CreateCell(i);
                cell.CellStyle = headerCellStyle;
                cell.SetCellValue(commentDictionary[dataSource.Columns[i].Caption][1]);
            }

            //数据行单元格的样式
            ICellStyle dataCellStyle = workbook.CreateCellStyle();
            dataCellStyle.VerticalAlignment = VerticalAlignment.Center;
            dataCellStyle.Alignment = HorizontalAlignment.Center;
            IFont fontBody = workbook.CreateFont();
            fontBody.Boldweight = (short)FontBoldWeight.Normal;
            fontBody.FontHeightInPoints = 11;
            dataCellStyle.SetFont(fontBody);

            for (Int32 i = 0; i < dataSource.Rows.Count; i++)
            {
                IRow row = sheet.CreateRow(i + 3);
                for (Int32 j = 0; j < dataSource.Columns.Count; j++)
                {
                    ICell cell = row.CreateCell(j);
                    cell.SetCellValue(dataSource.Rows[i][j].ToString());
                    cell.CellStyle = dataCellStyle;
                }
            }

            //单元格自适应宽度
            for (Int32 i = 0; i < dataSource.Columns.Count; i++)
            {
                sheet.AutoSizeColumn(i);
            }

            //将表单写入文件
            workbook.Write(File.Create(path));
        }

        #endregion
    }
}