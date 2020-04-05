using System;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace DocConvert_Server
{
    public partial class Form1 : Form
    {
        System.Windows.Threading.DispatcherTimer workProcessTimer = new System.Windows.Threading.DispatcherTimer();
        MainServer socketServer = new MainServer();

        public Form1()
        {
            InitializeComponent();
            CheckForIllegalCrossThreadCalls = false;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            #region Create SocketServer
            socketServer.InitConfig();
            socketServer.CreateServer();
            var IsResult = socketServer.Start();

            if (IsResult)
            {
                DevLog.Write("[Socket Server Listening...]", LOG_LEVEL.INFO);
                DevLog.Write(string.Format("[INFO] IP: {0}   포트: {1}   프로토콜: {2}   서버이름: {3}", Properties.Settings.Default.serverIP, socketServer.Config.Port, socketServer.Config.Mode, socketServer.Config.Name), LOG_LEVEL.INFO);

                pictureBox1.Image = DocConvert_Server.Properties.Resources.success_icon;
            }
            else
            {
                DevLog.Write(string.Format("[ERROR] 서버 네트워크 시작 실패"), LOG_LEVEL.ERROR);
                pictureBox1.Image = DocConvert_Server.Properties.Resources.error_icon;
                return;
            }

            workProcessTimer.Tick += new EventHandler(OnProcessTimedEvent);
            workProcessTimer.Interval = new TimeSpan(0, 0, 0, 0, 32);
            workProcessTimer.Start();
            new Thread(delegate ()
            {
                    while (!this.IsDisposed)
                    {
                        toolStripStatusLabel2.Text = string.Format("        Session Connect Count: {0}/{1}", socketServer.SessionCount, Properties.Settings.Default.socketSessionCount);
                        if (IsTcpPortAvailable(Properties.Settings.Default.fileServerPORT))
                            pictureBox2.Image = Properties.Resources.success_icon;
                        else
                            pictureBox2.Image = Properties.Resources.error_icon;
                        Thread.Sleep(1000);
                    }
            }).Start();
            #endregion
            #region Create File Server
            int fileServerPort = Properties.Settings.Default.fileServerPORT;

            #endregion
        }

        #region Socket Method

        private void OnProcessTimedEvent(object sender, EventArgs e)
        {
            int logWorkCount = 0;

            while (true)
            {
                string msg;

                if (DevLog.GetLog(out msg))
                {
                    ++logWorkCount;

                    toolStripStatusLabel1.Text = string.Format("LogCount: {0}/{1}", listBoxLog.Items.Count, Properties.Settings.Default.LogMaxCount);

                    if (listBoxLog.Items.Count >= Properties.Settings.Default.LogMaxCount)
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

        #endregion

        #region File Method
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

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            socketServer.Dispose();
            Application.Exit();
        }

        private void listBoxLog_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox1.Text = "";
            foreach(var logLIst in listBoxLog.SelectedItems)
            {
                textBox1.AppendText(logLIst.ToString() + "\r\n");
            }
        }
    }
}
