using FubarDev.FtpServer;
using FubarDev.FtpServer.FileSystem.DotNet;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Net.NetworkInformation;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DocConvert_FileServer
{
    public partial class Main : Form
    {
        System.Windows.Threading.DispatcherTimer workProcessTimer = new System.Windows.Threading.DispatcherTimer();
        public Main()
        {
            InitializeComponent();
            CheckForIllegalCrossThreadCalls = false;
        }

        private void Main_Load(object sender, EventArgs e)
        {
            workProcessTimer.Tick += new EventHandler(OnProcessTimedEvent);
            workProcessTimer.Interval = new TimeSpan(0, 0, 0, 0, 32);
            workProcessTimer.Start();

            new Thread(delegate ()
            {
                while (!this.IsDisposed)
                {
                    if (IsTcpPortAvailable(DocConvert_Server.getSettings.getSocketPORT))
                        pictureBox1.Image = Properties.Resources.success_icon;
                    else
                        pictureBox1.Image = Properties.Resources.error_icon;
                    Thread.Sleep(1000);
                }
            }).Start(); 
            new Thread(delegate ()
            {
                startServer();
            }).Start();
        }
        public async void startServer()
        {
            pictureBox2.Image = Properties.Resources.success_icon;
            await startFileServer();
        }

        #region File Server Method
        static async Task startFileServer()
        {
            // Setup dependency injection
            var services = new ServiceCollection();

            // use %TEMP%/TestFtpServer as root folder
            services.Configure<DotNetFileSystemOptions>(opt => opt
                .RootPath = DocConvert_Server.getSettings.getDataPath);

            // Add FTP server services
            // DotNetFileSystemProvider = Use the .NET file system functionality
            // AnonymousMembershipProvider = allow only anonymous logins
            services.AddFtpServer(builder => builder
                .UseDotNetFileSystem() // Use the .NET file system functionality
                .EnableAnonymousAuthentication()); // allow anonymous logins


            // Configure the FTP server
            services.Configure<FtpServerOptions>(opt => opt.ServerAddress = "*");
            services.Configure<FtpServerOptions>(opt => opt.Port = DocConvert_Server.getSettings.getFileServerPORT);

            // Build the service provider
            using (var serviceProvider = services.BuildServiceProvider())
            {
                // Initialize the FTP server
                var ftpServerHost = serviceProvider.GetRequiredService<IFtpServerHost>();

                // Start the FTP server
                await ftpServerHost.StartAsync();
                DevLog.Write("[File Server Listening...]", LOG_LEVEL.INFO);
                while (true)
                {
                    Thread.Sleep(3000000);
                }

                // Stop the FTP server
                /*await ftpServerHost.StopAsync();*/
            }
        }
        #endregion
        #region Socket Server Method
        public static bool IsTcpPortAvailable(int tcpPort)
        {
            var ipgp = IPGlobalProperties.GetIPGlobalProperties();

            // Check LISTENING ports
            IPEndPoint[] endpoints = ipgp.GetActiveTcpListeners();
            foreach (var ep in endpoints)
            {
                if (ep.Port == tcpPort)
                {
                    return true;
                }
            }

            return false;
        }
        #endregion

        private void listBoxLog_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox1.Text = "";
            foreach (var logLIst in listBoxLog.SelectedItems)
            {
                textBox1.AppendText(logLIst.ToString() + "\r\n");
            }
        }

        private void OnProcessTimedEvent(object sender, EventArgs e)
        {
            // 너무 이 작업만 할 수 없으므로 일정 작업 이상을 하면 일단 패스한다.
            int logWorkCount = 0;

            while (true)
            {
                string msg;

                if (DevLog.GetLog(out msg))
                {
                    ++logWorkCount;

                    toolStripStatusLabel1.Text = string.Format("LogCount: {0}/{1}", listBoxLog.Items.Count, DocConvert_Server.getSettings.getLongMaxCount);

                    if (listBoxLog.Items.Count >= DocConvert_Server.getSettings.getLongMaxCount)
                    {
                        listBoxLog.Items.RemoveAt(0);
                    }

                    listBoxLog.Items.Add(msg);
                }
                else
                {
                    break;
                }

                if (logWorkCount > 7)
                {
                    break;
                }
            }
        }

        private void Main_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.ExitThread();
            Environment.Exit(0);
        }
    }
}
