using NLog;
using System;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Windows.Forms;

using DocConvert_Server;

namespace DocConvert_Server
{
    public partial class Form1 : Form
    {
        System.Windows.Threading.DispatcherTimer workProcessTimer = new System.Windows.Threading.DispatcherTimer();
        public Form1()
        {
            InitializeComponent();
            CheckForIllegalCrossThreadCalls = false;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            #region Create SocketServer
            var server = new MainServer();
            server.InitConfig();
            server.CreateServer();

            var IsResult = server.Start();

            if (IsResult)
            {
                DevLog.Write(string.Format("[서비스 시작]"), LOG_LEVEL.INFO);
                DevLog.Write(string.Format("[INFO] IP: {0}   포트: {1}   프로토콜: {2}   서버이름: {3}", client_IP, server.Config.Port, server.Config.Mode, server.Config.Name), LOG_LEVEL.INFO);

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
            #endregion
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

        private string client_IP
        {
            get
            {
                IPHostEntry host = Dns.GetHostEntry(Dns.GetHostName());
                string ClientIP = string.Empty;
                for (int i = 0; i < host.AddressList.Length; i++)
                {
                    if (host.AddressList[i].AddressFamily == AddressFamily.InterNetwork)
                    {
                        ClientIP = host.AddressList[i].ToString();
                    }
                }
                return ClientIP;
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
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
