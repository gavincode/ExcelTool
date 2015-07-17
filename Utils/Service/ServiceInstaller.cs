using Microsoft.Win32;
using System;
using System.Collections;
using System.Configuration.Install;
using System.ServiceProcess;

namespace SierviceManager.Install
{
    /// <summary>
    /// 服务安装类
    /// </summary>
    public class ServiceInstaller
    {
        #region 检查服务存在的存在性

        /// <summary>
        /// 检查服务存在的存在性
        /// </summary>
        /// <param name="serviceName">服务名</param>
        /// <returns>存在返回 true,否则返回 false;</returns>
        public static Boolean IsServiceIsExisted(String serviceName)
        {
            ServiceController[] services = ServiceController.GetServices();

            foreach (ServiceController controller in services)
            {
                if (String.Equals(controller.ServiceName, serviceName, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        #endregion

        #region 安装Windows服务

        /// <summary>
        /// 安装Windows服务
        /// </summary>
        /// <param name="stateSaver">状态集合</param>
        /// <param name="filepath">程序文件路径</param>
        public static void InstallService(IDictionary stateSaver, String filepath)
        {
            AssemblyInstaller AssemblyInstaller1 = new AssemblyInstaller();
            AssemblyInstaller1.UseNewContext = true;
            AssemblyInstaller1.Path = filepath;
            AssemblyInstaller1.Install(stateSaver);
            AssemblyInstaller1.Commit(stateSaver);
            AssemblyInstaller1.Dispose();
        }

        #endregion

        #region 卸载Windows服务

        /// <summary>
        /// 卸载Windows服务
        /// </summary>
        /// <param name="filepath">程序文件路径</param>
        public static void UnInstallService(String filepath)
        {
            AssemblyInstaller AssemblyInstaller1 = new AssemblyInstaller();
            AssemblyInstaller1.UseNewContext = true;
            AssemblyInstaller1.Path = filepath;
            AssemblyInstaller1.Uninstall(null);
            AssemblyInstaller1.Dispose();
        }

        #endregion

        #region 判断window服务是否启动

        /// <summary>
        /// 判断某个Windows服务是否启动
        /// </summary>
        /// <param name="serviceName">服务名</param>
        /// <returns>已启动true;否则false</returns>
        public static Boolean IsServiceStart(String serviceName)
        {
            ServiceController controller = new ServiceController(serviceName);

            if (!controller.Status.Equals(ServiceControllerStatus.Stopped))
            {
                return true;
            }

            return false;
        }

        #endregion

        #region  修改服务的启动项

        /// <summary>  
        /// 修改服务的启动项 2为自动,3为手动  
        /// </summary>  
        /// <param name="startType">开机启动方式: 2为自动,3为手动  </param>  
        /// <param name="serviceName">服务名</param>  
        /// <returns>成功true;否则false</returns>  
        public static Boolean ChangeServiceStartType(Int32 startType, String serviceName)
        {
            try
            {
                RegistryKey regist = Registry.LocalMachine;
                RegistryKey sysReg = regist.OpenSubKey("SYSTEM");
                RegistryKey currentControlSet = sysReg.OpenSubKey("CurrentControlSet");
                RegistryKey services = currentControlSet.OpenSubKey("Services");
                RegistryKey servicesName = services.OpenSubKey(serviceName, true);
                servicesName.SetValue("Start", startType);
            }
            catch (Exception)
            {
                return false;
            }

            return true;
        }

        #endregion

        #region 启动服务

        /// <summary>
        /// 启动服务
        /// </summary>
        /// <param name="serviceName">服务名</param>
        /// <param name="maxWaitSecond">最大等待时间/秒 默认30秒</param>
        /// <returns>启动成功true;否则false</returns>
        public static Boolean StartService(String serviceName, Int32 maxWaitSecond = 30)
        {
            //判断服务是否存在
            if (IsServiceIsExisted(serviceName))
            {
                ServiceController controller = new ServiceController(serviceName);

                if (controller.Status != ServiceControllerStatus.Running && controller.Status != ServiceControllerStatus.StartPending)
                {
                    controller.Start();
                    controller.WaitForStatus(ServiceControllerStatus.Running, new TimeSpan(0, 0, 0, maxWaitSecond));
                    controller.Close();
                    return true;
                }
            }

            return false;
        }

        #endregion

        #region 停止服务

        /// <summary>
        /// 停止服务
        /// </summary>
        /// <param name="serviceName">服务名</param>
        /// <param name="maxWaitSecond">最大等待时间/秒 默认30秒</param>
        /// <returns>成功true;否则false</returns>
        public static Boolean StopService(String serviceName, Int32 maxWaitSecond = 30)
        {
            //检验服务是否存在
            if (IsServiceIsExisted(serviceName))
            {
                ServiceController service = new ServiceController(serviceName);

                if (service.Status == ServiceControllerStatus.Running)
                {
                    service.Stop();
                    service.WaitForStatus(ServiceControllerStatus.Stopped, new TimeSpan(0, 0, 0, maxWaitSecond));
                    service.Close();

                    return true;
                }
            }

            return false;
        }

        #endregion
    }
}