using System.ComponentModel;
using System;
using System.Linq;

namespace Service
{
    using Moqikaka.Util;
    using Utils.Configuration;
    using System.IO;
    using System.Diagnostics;
    using Utils.Log;
    using System.Reflection;

    [RunInstaller(true)]
    public partial class ProjectInstaller : System.Configuration.Install.Installer
    {
        /// <summary>
        /// 构造函数
        /// </summary>
        public ProjectInstaller()
        {
            InitializeComponent();

            this.serviceInstaller1.ServiceName = ConfigurationHelper.AppSettings["ServiceName"];
        }

        /// <summary>
        /// 静态构造函数
        /// </summary>
        static ProjectInstaller()
        {
            String assemblyPath = Directory.GetParent(Assembly.GetExecutingAssembly().Location).FullName;

            //初始化app.config路径
            ConfigurationHelper.ConfigFile = Assembly.GetExecutingAssembly().Location + ".config";
            ConfigurationHelper.Init();

            //EventLog.WriteEntry("ProjectInstaller", ConfigurationHelper.ConfigFile);

            //初始化路径
            Trace.Listeners.Clear();  //清除系统监听器 (就是输出到Console的那个)
            Trace.Listeners.Add(new LogTrace(Path.Combine(assemblyPath, "Log"))); //添加LogTrace实例

            //初始化监听目录
            String listenFolder = Path.Combine(assemblyPath, ConfigurationHelper.AppSettings["ExcelFolder"]);

            if (!Directory.Exists(listenFolder))
                Directory.CreateDirectory(listenFolder);

            //初始化默认文件夹
            foreach (var item in ConfigurationHelper.AppSettings.AllSettings.Where(p => p.Value.StartsWith("server")))
            {
                String folder = Path.Combine(listenFolder, item.Key);

                if (!Directory.Exists(folder))
                {
                    Directory.CreateDirectory(folder);
                    Directory.CreateDirectory(Path.Combine(folder, "2015-01-01 01.00(sample)"));
                }
            }
        }

        /// <summary>
        /// 安装
        /// </summary>
        /// <param name="stateSaver"></param>
        public override void Install(System.Collections.IDictionary stateSaver)
        {
            base.Install(stateSaver);

            Trace.Write("服务安装成功");
        }

        public override void Uninstall(System.Collections.IDictionary savedState)
        {
            base.Uninstall(savedState);

            Trace.Write("服务卸载成功");
        }
    }
}
