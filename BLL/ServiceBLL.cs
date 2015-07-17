// ****************************************
// FileName:ServiceBLL.cs
// Description:Excel自动导入服务逻辑处理类
// Tables:Nothing
// Author:Gavin
// Create Date:2014/11/27 14:46:04
// Revision History:
// ****************************************

using System;
using System.Data;
using System.Linq;
using System.Collections.Generic;
using System.IO;
using System.Diagnostics;
using System.Text;
using Utils.Excel;
using System.Threading;
using Utils.Configuration;
using Utils.Log;
using System.Globalization;
using System.Configuration;

namespace BLL
{
    /// <summary>
    /// Excel自动导入服务逻辑处理类
    /// </summary>
    public static class ServiceBLL
    {
        #region 静态变量及初始化

        //监听Excel文件夹
        private static String ExcelFolder = String.Empty;

        //是否正在导入
        private static Boolean IsImporting = false;

        //导入结果
        private static Dictionary<String, Int32> resultInfo = new Dictionary<String, Int32>();

        //常用时间字符串格式集合
        public static List<String> DateTimeFormatList = new List<String>();

        /// <summary>
        /// 静态构造函数
        /// </summary>
        static ServiceBLL()
        {
            //初始化app.config路径
            ConfigurationHelper.ConfigFile = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None).FilePath;
            ConfigurationHelper.Init();

            //初始化log日志路径
            Trace.Listeners.Clear();  //清除系统监听器 (就是输出到Console的那个)
            Trace.Listeners.Add(new LogTrace(AppDomain.CurrentDomain.BaseDirectory + "\\Log")); //添加LogTrace实例

            //初始化监听文件夹
            ExcelFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ConfigurationHelper.AppSettings["ExcelFolder"]);
            if (!Directory.Exists(ExcelFolder))
                Directory.CreateDirectory(ExcelFolder);

            DateTimeFormatList.AddRange(new List<String>{
            "yyyy年MM月dd日HH时mm分ss秒",
            "yyyy年MM月dd日HH时mm分",
            "yyyy年MM月dd日HH时",
            "yyyy-MM-dd HH.mm.ss",
            "yyyy-MM-dd HH.mm",
            "yyyy-MM-dd HH:mm:ss",
            "yyyy-MM-dd HH:mm",
            "yyyy-MM-dd HH", 
            "yyyyMMddHHmmss",
            "yyyyMMddHHmm",
            "yyyyMMddHH",
            "yyyyMMdd"});
        }

        #endregion

        /// <summary>
        /// 开始监听
        /// </summary>
        public static void Begin()
        {
            try
            {
                //监听主线程
                Thread listenThread = new Thread(new ThreadStart(Listen));
                listenThread.Start();

                Trace.Write("服务启动成功");
            }
            catch (Exception ex)
            {
                Trace.Write("错误信息: " + ex.Message + "\r\n" + ex.StackTrace);
            }
        }

        /// <summary>
        /// 监听
        /// </summary>
        private static void Listen()
        {
            try
            {
                //1. 无限循环, 每分钟监听一次文件夹
                while (true)
                {
                    //如果当前正在导入,休眠后继续尝试
                    if (IsImporting)
                    {
                        Thread.Sleep((60 - DateTime.Now.Second) * 1000);
                        continue;
                    }

                    try
                    {
                        //收集当前需要导入的文件夹列表
                        var importFolderList = CollectFolders();
                        if (importFolderList.Count > 0)
                        {
                            //异步批量导入
                            Action<List<String>> action = new Action<List<String>>(ImprotFolders);
                            action.BeginInvoke(importFolderList, CallBack, importFolderList);
                        }
                    }
                    catch (Exception ec) //导入过程发生异常,记录日志后继续监听
                    {
                        Trace.Write("错误信息: " + ec.Message + "\r\n" + ec.StackTrace);
                        continue;
                    }

                    Thread.Sleep((60 - DateTime.Now.Second) * 1000);
                }
            }
            catch (Exception ex)
            {
                Trace.Write("错误信息: " + ex.Message + "\r\n" + ex.StackTrace);
            }
            finally
            {
                Trace.Write("服务已经停止");
            }
        }

        /// <summary>
        /// 收集当前需要导入的文件夹列表
        /// </summary>
        private static List<String> CollectFolders()
        {
            //需要导入的文件夹列表
            List<String> folderList = new List<String>();

            DateTime importTime = DateTime.MinValue;
            DateTime now = DateTime.Now;

            //遍历所有文件夹,若文件夹名称时间等于当前时间
            foreach (var folderName in Directory.GetDirectories(ExcelFolder)) //遍历监听主目录
            {
                foreach (var folder in Directory.GetDirectories(folderName))  //遍历各个数据库导入目录
                {
                    //文件夹时间是否满足
                    if (Path.GetFileName(folder).IsDateTime(ref importTime)
                        && importTime.Date == now.Date
                        && importTime.Hour == now.Hour
                        && importTime.Minute == now.Minute)
                    {
                        folderList.Add(folder);
                    }
                }
            }

            return folderList;
        }

        /// <summary>
        /// 导入指定文件夹列表
        /// </summary>
        /// <param name="folderList">指定文件夹列表</param>
        private static void ImprotFolders(List<String> folderList)
        {
            IsImporting = true;

            var dbSettings = ConfigurationHelper.AppSettings.AllSettings;

            //循环导入文件夹
            foreach (var folder in folderList)
            {
                Trace.Write("开始导入文件夹: " + Path.GetFullPath(folder));
                resultInfo.Clear();

                //当前需要导入的数据库连接字符串
                String dbKey = new FileInfo(folder).Directory.Name;

                //监听数据库关键字不存在
                if (!dbSettings.ContainsKey(dbKey))
                {
                    Trace.Write(String.Format("监听目录: {0}, 对应的数据库连接配置不存在!", dbKey));
                    continue;
                }

                //设置数据连接字符串
                ExcelBLL.SetDbConnection(dbSettings[dbKey]);

                //Check数据库连接是否正确
                if (!ExcelBLL.IsDataBaseAccess())
                {
                    Trace.Write(String.Format("监听目录: {0}, 对应的数据库连接配置不正确!", dbKey));
                    continue;
                }

                //收集导入excel文件
                var excelFiles = Directory.GetFiles(folder).Where(p => p.EndsWith(".xlsx") && !p.Contains("~"));

                //同步导入
                BatchImport(excelFiles);

                //导入完成
                CallBack(folder);
            }

            IsImporting = false;
        }

        /// <summary>
        /// 批量导入excel文件
        /// </summary>
        /// <param name="fileArray">需要导入的文件集合</param>
        private static void BatchImport(IEnumerable<String> fileArray)
        {
            //获取数据库表名集合
            var dbTableList = ExcelBLL.GetTableNameList();

            //重新读取忽略表单
            IgnoreSheetsBLL.Reset();

            //遍历导入每个Excel文档
            foreach (var fileName in fileArray)
            {
                //导入单个excel文件
                ImportExcel(fileName, dbTableList);
            }
        }

        /// <summary>
        /// 导入单个excel文件
        /// </summary>
        /// <param name="path">excel文档路径</param>
        /// <param name="dbTables">数据库表名集合</param>
        private static void ImportExcel(String path, List<String> dbTables)
        {
            MoqikakaExcel excel = new MoqikakaExcel(path);

            //遍历导入Excel每个表单
            for (int i = 0; i < excel.NumberOfSheets; i++)
            {
                //导入单个表单
                ImportSheet(excel, i, dbTables);
            }
        }

        /// <summary>
        /// 导入excel单个表单
        /// </summary>
        /// <param name="excel">MoqikakaExcel对象</param>
        /// <param name="sheetIndex">导入表单序号</param>
        /// <param name="dbTables">数据库表名集合</param>
        private static void ImportSheet(MoqikakaExcel excel, Int32 sheetIndex, List<String> dbTables)
        {
            String sheetName = excel.GetSheetName(sheetIndex);

            //排除没用的表单/已忽略表单/数据库不存在的表单
            if (ExcelBLL.IsUselessSheet(sheetName)
                || IgnoreSheetsBLL.IsIgnoreSheet(sheetName)
                || !dbTables.Contains(sheetName.ToLower()))
                return;

            //读取表单数据
            DataTable table = ExcelBLL.TryRead(excel, sheetIndex);
            if (table == null)
            {
                resultInfo[sheetName] = 0;
                return;
            }

            try
            {
                List<String> sqlList = ExcelBLL.GetSQL(table);

                Int32 rows = ExcelBLL.Insert(sqlList, table.TableName, excel.Path, false);

                resultInfo[table.TableName] = rows;
            }
            catch (Exception ex)
            {
                resultInfo[sheetName] = 0;
                Trace.Write("错误表单: " + sheetName + "\r\n" + "错误信息: " + ex.Message + "\r\n" + ex.StackTrace);
            }
        }

        /// <summary>
        /// 完成回调
        /// </summary>
        /// <param name="result"></param>
        private static void CallBack(IAsyncResult result)
        {
            IsImporting = false;
        }

        /// <summary>
        /// 完成回调
        /// </summary>
        /// <param name="folder">文件夹</param>
        private static void CallBack(String folder)
        {
            Boolean logError = false;
            StringBuilder builder = new StringBuilder();
            StringBuilder errorBuilder = new StringBuilder();

            builder.AppendLine("导入时间: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            builder.AppendLine("本次导入明细如下:");

            errorBuilder.AppendLine("本次导入异常表单如下:");

            foreach (var item in resultInfo)
            {
                builder.AppendLine(String.Format("{0} 导入数量为: {1}", item.Key, item.Value));

                if (item.Value == 0)
                {
                    logError = true;
                    errorBuilder.AppendLine(String.Format("{0} 导入数量为: {1}", item.Key, item.Value));
                }
            }

            Trace.Write(builder.ToString());

            if (logError)
                Trace.Write(errorBuilder.ToString());

            //重命名导入文件夹
            if (logError || resultInfo.Count == 0)
                AppendName(folder, "Error");
            else
                AppendName(folder, "Done");
        }

        /// <summary>
        /// 追加文件名
        /// </summary>
        /// <param name="folder">文件名</param>
        /// <param name="appendName">追加的名称</param>
        private static void AppendName(String folder, String appendName)
        {
            DirectoryInfo info = new DirectoryInfo(folder);

            info.MoveTo(info.FullName.Replace(info.Name, info.Name + "&" + appendName));
        }

        /// <summary>
        /// 转换字符串为日期时间.如果转换失败,则返回指定的默认值
        /// </summary>
        /// <param name="value">要转换的字符串</param>
        /// <param name="dateTime">dateTime</param>
        /// <returns>是否为DateTime</returns>
        public static Boolean IsDateTime(this string value, ref DateTime dateTime)
        {
            foreach (var format in DateTimeFormatList)
            {
                if (DateTime.TryParseExact(value, format, null, DateTimeStyles.None, out dateTime))
                    return true;
            }

            return false;
        }
    }
}