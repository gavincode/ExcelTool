// ****************************************
// FileName:LogTrace.cs
// Description:日志记录类
// Tables:None
// Author:Gavin && Burney
// Create Date:2014-06-01
// Revision History:
// ****************************************

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace Utils.Log
{
    public class LogTrace : TraceListener
    {
        //日志文件夹
        private static String logFolder = String.Empty;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="folder"></param>
        public LogTrace(String folder)
        {
            logFolder = folder;

            if (!Directory.Exists(logFolder))
                Directory.CreateDirectory(logFolder);
        }

        /// <summary>
        /// 日志文件
        /// </summary>
        public static String LogTxt
        {
            get
            {
                return String.Format("{0}\\log{1}.txt", logFolder, DateTime.Now.ToString("yyyy-MM-dd"));
            }
        }

        /// <summary>
        /// Write方法的日志格式
        /// </summary>
        private String _messageFormater
        {
            get
            {
                String beginTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("***********************" + beginTime + "****************************");
                sb.AppendLine();
                sb.AppendLine("{0}");
                sb.AppendLine();
                sb.AppendLine("**********************************************************************");
                sb.AppendLine();
                sb.AppendLine();
                sb.AppendLine();
                return sb.ToString();
            }
        }

        /// <summary>
        /// 已固定格式记录日志
        /// </summary>
        /// <param name="message">日志信息</param>
        public override void Write(String message)
        {
            if (!File.Exists(LogTxt))
                File.CreateText(LogTxt).Close();

            File.AppendAllText(LogTxt, String.Format(_messageFormater, message));
        }

        /// <summary>
        /// 记录日志行
        /// </summary>
        /// <param name="o"></param>
        public override void WriteLine(String o)
        {
        }

        /// <summary>
        /// Write扩展方法,记录错误信息
        /// </summary>
        /// <param name="message">日志信息</param>
        /// <param name="category">含 sql 或 error 字符串的文件路径</param>
        public override void Write(String message, String category)
        {
            if (category.ToLower().IndexOf("sql", StringComparison.Ordinal) != -1)
            {
                File.AppendAllText(category, message);
                return;
            }

            if (category.ToLower().IndexOf("error", StringComparison.Ordinal) != -1)
            {
                File.AppendAllText(LogTxt.Replace("log", "error"), String.Format(_messageFormater, message));
                return;
            }
        }
    }
}
