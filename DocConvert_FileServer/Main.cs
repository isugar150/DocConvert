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
            DevLog.Write("테스트", LOG_LEVEL.INFO);
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
        }

        #region File Server Method

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
    }
}
