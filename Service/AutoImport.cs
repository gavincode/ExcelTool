using BLL;
using System.Diagnostics;
using System.ServiceProcess;

namespace Service
{
    public partial class AutoImport : ServiceBase
    {
        public AutoImport()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            ServiceBLL.Begin();
        }

        protected override void OnStop()
        {
            base.OnStop();

            Trace.Write("服务已停止");
        }
    }
}
