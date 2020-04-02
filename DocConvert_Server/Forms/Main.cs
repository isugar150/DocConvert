﻿using NLog;
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
        private static Logger logger = LogManager.GetLogger("DocConvert_Server_Log");
        public Form1()
        {
            InitializeComponent();
            CheckForIllegalCrossThreadCalls = false;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            var server = new MainServer();
            server.InitConfig();
            server.CreateServer();

            var IsResult = server.Start();

            if (IsResult)
            {
                DevLog.Write(string.Format("서버 네트워크 시작"), LOG_LEVEL.INFO);
            }
            else
            {
                DevLog.Write(string.Format("[ERROR] 서버 네트워크 시작 실패"), LOG_LEVEL.ERROR);
                return;
            }

            workProcessTimer.Tick += new EventHandler(OnProcessTimedEvent);
            workProcessTimer.Interval = new TimeSpan(0, 0, 0, 0, 32);
            workProcessTimer.Start();
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

                    if (listBoxLog.Items.Count > 100)
                    {
                        listBoxLog.Items.Clear();
                    }

                    listBoxLog.Items.Add(msg);
                    listBoxLog.SelectedIndex = listBoxLog.Items.Count - 1;
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

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }

        private void consoleAppend(string text)
        {
            listBoxLog.Items.Add(System.DateTime.Now.ToString("yyyy/MM/dd hh:mm:ss.fff") + "   " + text + "\r\n");
        }
    }
}
