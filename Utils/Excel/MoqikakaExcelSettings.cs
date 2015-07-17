// ****************************************
// FileName:MoqikakaExcelSettings
// Description: 摩奇卡卡Excel模板 静态设置读取类
// Tables:
// Author:Gavin
// Create Date:2014/6/6 9:38:40
// Revision History:
// ****************************************

using System;
using System.Collections.Generic;

namespace Utils.Excel
{
    using Utils.Configuration;

    /// <summary>
    /// 摩奇卡卡导入Excel中,特殊行设置信息
    /// </summary>
    public static class MoqikakaExcelSettings
    {
        #region 初始化

        /// <summary>
        /// 初始化
        /// </summary>
        public static void Init()
        {
            Int32.TryParse(ConfigurationHelper.AppSettings["DataColumnNameDescRowNum"], out DescRowNum);
            Int32.TryParse(ConfigurationHelper.AppSettings["DataColumnNameRowNum"], out NameRowNum);
            Int32.TryParse(ConfigurationHelper.AppSettings["DataTypeRowNum"], out TypeRowNum);
            Int32.TryParse(ConfigurationHelper.AppSettings["DataRowNum"], out DataRowNum);

            if (DescRowNum != -1) SpecialRowList.Add(DescRowNum);
            if (NameRowNum != -1) SpecialRowList.Add(NameRowNum);
            if (TypeRowNum != -1) SpecialRowList.Add(TypeRowNum);
        }

        /// <summary>
        /// 静态构造函数
        /// </summary>
        static MoqikakaExcelSettings()
        {
            Init();
        }

        #endregion

        /// <summary>
        /// 中文描述行，若无，则设为-1
        /// </summary>
        public static Int32 DescRowNum = -1;

        /// <summary>
        /// 数据库字段名行;若无，则设为-1
        /// </summary>
        public static Int32 NameRowNum = -1;

        /// <summary>
        /// 数据库字段类型行;若无，则设为-1
        /// </summary>
        public static Int32 TypeRowNum = -1;

        /// <summary>
        /// 数据开始行【必须设置，且大于0】
        /// </summary>
        public static Int32 DataRowNum = -1;

        /// <summary>
        /// 有值的特殊行行号列表
        /// </summary>
        public static List<Int32> SpecialRowList = new List<Int32>();
    }
}
