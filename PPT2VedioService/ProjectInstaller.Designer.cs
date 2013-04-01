namespace PPT2VedioService
{
    partial class ProjectInstaller
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.spiMain = new System.ServiceProcess.ServiceProcessInstaller();
            this.siMain = new System.ServiceProcess.ServiceInstaller();
            // 
            // spiMain
            // 
            this.spiMain.Account = System.ServiceProcess.ServiceAccount.LocalSystem;
            this.spiMain.Password = null;
            this.spiMain.Username = null;
            // 
            // siMain
            // 
            this.siMain.Description = "PPT转换视频服务";
            this.siMain.DisplayName = "PPT2Vedio";
            this.siMain.ServiceName = "PPT2Vedio";
            // 
            // ProjectInstaller
            // 
            this.Installers.AddRange(new System.Configuration.Install.Installer[] {
            this.spiMain,
            this.siMain});

        }

        #endregion

        private System.ServiceProcess.ServiceProcessInstaller spiMain;
        private System.ServiceProcess.ServiceInstaller siMain;
    }
}