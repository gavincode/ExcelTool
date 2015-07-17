// ****************************************
// FileName:Program.cs
// Description:程序入口类
// Tables:None
// Author:Gavin && Burney
// Create Date:2014-06-01
// Revision History:
// ****************************************

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;

namespace ExcelTool
{
    using System.IO;
    using Utils.Log;
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            //输出日志替换为自定义模式
            Trace.Listeners.Clear();
            Trace.Listeners.Add(new LogTrace(Environment.CurrentDirectory + "\\Log"));

            //初始化SQL存放文件路径
            if (!Directory.Exists(Environment.CurrentDirectory + "\\SQL"))
                Directory.CreateDirectory(Environment.CurrentDirectory + "\\SQL");

            Application.Run(new Main());
        }
    }
}
