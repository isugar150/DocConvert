using NLog;
using NLog.Internal;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

using DocConvert_Core.IniLib;
using System.IO;
using DocConvert_Core.interfaces;

namespace DocConvert_Manager
{
    public partial class Manager : Form
    {
        private System.Windows.Threading.DispatcherTimer workProcessTimer = new System.Windows.Threading.DispatcherTimer();
        public static iniProperties IniProperties = new iniProperties();
        private static System.Windows.Forms.Timer tScheduler;
        private Point mousePoint;
        private int targetPID = 0;
        public Manager()
        {
            InitializeComponent();
            CheckForIllegalCrossThreadCalls = false;
            #region 로깅 서비스 등록
            workProcessTimer.Tick += new EventHandler(OnProcessTimedEvent);
            workProcessTimer.Interval = new TimeSpan(0, 0, 0, 0, 32);
            workProcessTimer.Start();
            #endregion
            #region 환경 설정 파싱
            IniFile pairs = new IniFile();
            while (true)
            {
                try
                {
                    if (new FileInfo("./DocConvert_Manager.ini").Exists)
                    {
                        pairs.Load("./DocConvert_Manager.ini");
                        IniProperties.targetPath = pairs["DC Manager"]["targetPath"].ToString();
                        IniProperties.minimized = pairs["DC Manager"]["minimized"].ToString().Equals("Y");
                        IniProperties.refreshCycle = int.Parse(pairs["DC Manager"]["refreshCycle"].ToString());
                        break;
                    }
                    else
                    {
                        Setting.createSetting();
                        MessageBox.Show("설정 파일을 생성하였습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                catch (Exception)
                {
                    if (MessageBox.Show("설정 파일이 손상되었습니다. 초기화하시겠습니까?", "알림", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                    {
                        Setting.createSetting();
                        MessageBox.Show("설정 파일을 초기화하였습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                        program_Exit(true);
                }
            }
            #endregion
            #region 스케줄러 관련
            tScheduler = new System.Windows.Forms.Timer();
            tScheduler.Interval = CalculateTimerInterval();
            tScheduler.Tick += new EventHandler(tScheduler_Tick);
            tScheduler.Start();
            #endregion
        }

        private void Manager_Load(object sender, EventArgs e)
        {
            new Thread(delegate ()
            {
                while (!this.IsDisposed)
                {
                    try
                    {
                        Process[] targetProc = Process.GetProcessesByName("DocConvert_Server");
                        if (targetProc.Length != 0)
                        {
                            targetPID = targetProc[0].Id;
                            label3.Text = string.Format("PID: {0}", targetProc[0].Id);
                        }
                        else
                        {
                            targetPID = 0;
                            label3.Text = string.Format("No Detected");
                        }
                        if (targetPID == 0)
                        {
                            if (IniProperties.targetPath.Contains("./"))
                            {
                                if (IniProperties.minimized)
                                    Process.Start(Application.StartupPath + @"\" + IniProperties.targetPath.Replace("./", ""), "Minimized");
                                else
                                    Process.Start(Application.StartupPath + @"\" + IniProperties.targetPath.Replace("./", ""));
                            }
                            else
                            {
                                if (IniProperties.minimized)
                                    Process.Start(IniProperties.targetPath, "Minimized");
                                else
                                    Process.Start(IniProperties.targetPath);
                            }
                            DevLog.Write("타겟 프로세스를 실행하였습니다.");
                        }
                    }
                    catch (Exception e1)
                    {
                        DevLog.Write(e1.Message, LOG_LEVEL.ERROR);
                    }
                    Thread.Sleep(IniProperties.refreshCycle);
                }
            }).Start();
        }

        private void Manager_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
            program_Exit(false);
        }

        #region 스케줄러 메소드
        private int CalculateTimerInterval()
        {
            DateTime timeTaken = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 0, 0, 0).AddDays(1);

            TimeSpan curTime = timeTaken - DateTime.Now;

            Debug.WriteLine(timeTaken.ToString("yyyy-MM-dd HH:mm:ss"));
            Debug.WriteLine(curTime);
            return (int)curTime.TotalMilliseconds;
        }

        private void tScheduler_Tick(object sender, EventArgs e)
        {
            Thread.Sleep(3000);
            try
            {
                tScheduler.Interval = CalculateTimerInterval();
            }
            catch (Exception) { tScheduler.Dispose(); DevLog.Write("스케줄러 작동중 오류가 발생하여 비활성화 하였습니다."); }
            textBox1.Text = "";
        }

        #endregion

        #region 로깅 관련
        private void OnProcessTimedEvent(object sender, EventArgs e)
        {
            int logWorkCount = 0;

            while (true)
            {
                string msg;

                if (DevLog.GetLog(out msg))
                {
                    textBox1.AppendText(string.Format("{0}\r\n", msg));
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

        #region panel1 마우스 드래그 이벤트
        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            mousePoint = new Point(e.X, e.Y);
        }

        private void panel1_MouseMove(object sender, MouseEventArgs e)
        {
            if ((e.Button & MouseButtons.Left) == MouseButtons.Left)
            {
                Location = new Point(this.Left - (mousePoint.X - e.X),
                    this.Top - (mousePoint.Y - e.Y));
            }
        }
        #endregion

        private void program_Exit(bool forceExit)
        {
            if (!forceExit)
            {
                if (MessageBox.Show(this, "DC Manager를 종료하시겠습니까?", "경고", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK)
                {
                    try
                    {
                        Application.ExitThread();
                        Environment.Exit(0);
                    }
                    catch (Exception) { }
                }
            }
            else
            {
                try
                {
                    Application.ExitThread();
                    Environment.Exit(0);
                }
                catch (Exception) { }
            }
        }

        private void 프로그램종료ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            program_Exit(false);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Visible = false;
        }
    }
}
