using DocConvert_Server.Forms;
using DocConvert_Server.License;
using Microsoft.Win32;
using Newtonsoft.Json.Linq;
using NLog;
using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.NetworkInformation;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using vtortola.WebSockets;
using vtortola.WebSockets.Deflate;

// TODO FTPS 지원 (클라이언트단)

namespace DocConvert_Server
{
    public partial class Form1 : Form
    {
        private System.Windows.Threading.DispatcherTimer workProcessTimer = new System.Windows.Threading.DispatcherTimer();
        private MainServer socketServer = new MainServer();
        private int wsSessionCount = 0;
        private static System.Windows.Forms.Timer tScheduler;
        private const int CHECK_INTERVAL = 60; // 스케줄러 주기 (분)
        private WebSocketListener webSocketServer = null;
        private JObject checkLicense = new JObject();
        public static bool isHwpConverting = false;

        private static Logger logger = LogManager.GetLogger("DocConvert_Server_Log");

        public Form1(string[] args)
        {
            InitializeComponent();
            CheckForIllegalCrossThreadCalls = false;

            if (args.Length == 0)
            {
                DevLog.Write("[INFO] DocConvert Server를 옵션없이 실행하였습니다.", LOG_LEVEL.INFO);
            }
            else
            {
                string arg = "";
                for (int i = 0; i < args.Length; i++)
                {
                    arg += string.Format("   [{0}]: {1}", i, args[i].ToString());
                }
                DevLog.Write("[INFO] 아규먼트: " + arg, LOG_LEVEL.INFO);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.Text += " - " + Properties.Settings.Default.서버이름;
            textBox1.Text = ("┏━━━┓╋╋╋╋╋┏━━━┓╋╋╋╋╋╋╋╋╋╋╋╋╋┏┓╋┏━━━┓\r\n" +
                             "┗┓┏┓┃╋╋╋╋╋┃┏━┓┃╋╋╋╋╋╋╋╋╋╋╋╋┏┛┗┓┃┏━┓┃\r\n" +
                             "╋┃┃┃┣━━┳━━┫┃╋┗╋━━┳━┓┏┓┏┳━━┳┻┓┏┛┃┗━━┳━━┳━┳┓┏┳━━┳━┓\r\n" +
                             "╋┃┃┃┃┏┓┃┏━┫┃╋┏┫┏┓┃┏┓┫┗┛┃┃━┫┏┫┃╋┗━━┓┃┃━┫┏┫┗┛┃┃━┫┏┛\r\n" +
                             "┏┛┗┛┃┗┛┃┗━┫┗━┛┃┗┛┃┃┃┣┓┏┫┃━┫┃┃┗┓┃┗━┛┃┃━┫┃┗┓┏┫┃━┫┃\r\n" +
                             "┗━━━┻━━┻━━┻━━━┻━━┻┛┗┛┗┛┗━━┻┛┗━┛┗━━━┻━━┻┛╋┗┛┗━━┻┛\r\n");
            #region checkLicense
            try
            {
                Debug.WriteLine(Properties.Settings.Default.라이센스키);
                string licenseInfo = LicenseInfo.decryptAES256(Properties.Settings.Default.라이센스키, "JmDoCOnVerTerServErJmCoRp");
                checkLicense = JObject.Parse(licenseInfo);
                if (!checkLicense["HWID"].ToString().Equals(new LicenseInfo().getHWID())) { new MessageDialog("라이센스 오류", "라이센스 확인 후 다시시도하세요.", "HWID: " + new LicenseInfo().getHWID()).ShowDialog(this); program_Exit(true); return; }
                if (DateTime.Parse(checkLicense["EndDate"].ToString()) < DateTime.Now) { new MessageDialog("라이센스 오류", "라이센스 날짜가 만료되었습니다. 갱신후 다시시도해주세요.", "HWID: " + new LicenseInfo().getHWID()).ShowDialog(this); program_Exit(true); return; }

                DevLog.Write("[INFO] 나의 하드웨어 ID: " + new LicenseInfo().getHWID(), LOG_LEVEL.INFO);
                DevLog.Write(string.Format("[INFO] 라이센스 만료날짜: {0}", checkLicense["EndDate"].ToString()), LOG_LEVEL.INFO);
            }
            catch (Exception) { new MessageDialog("라이센스 오류", "라이센스키 파싱오류.", "HWID: " + new LicenseInfo().getHWID()).ShowDialog(this); program_Exit(true); return; }
            #endregion
            DirectoryInfo directoryInfo = new DirectoryInfo(Properties.Settings.Default.데이터경로 + @"\tmp");
            if (!directoryInfo.Exists)
                directoryInfo.Create();
            toolStripStatusLabel4.Text = "IP Address: " + Properties.Settings.Default.바인딩_IP;
            toolStripStatusLabel5.Text = "Socket Port: : " + Properties.Settings.Default.소켓서버포트.ToString();
            toolStripStatusLabel6.Text = "WebSocket Port: " + Properties.Settings.Default.웹소켓포트.ToString();
            toolStripStatusLabel7.Text = "File Server Port: " + Properties.Settings.Default.파일서버포트.ToString();

            checkBox1.Checked = Properties.Settings.Default.FollowTail;
            #region 한글 DLL 레지스트리 등록
            if (File.Exists(Application.StartupPath + @"\FilePathCheckerModuleExample.dll"))
            {
                RegistryKey regKey = Registry.CurrentUser.CreateSubKey(@"Software\HNC\HwpCtrl\Modules", RegistryKeyPermissionCheck.ReadWriteSubTree);
                DevLog.Write("[INFO] 한글 DLL을 레지스트리에 등록하였습니다.", LOG_LEVEL.INFO);
            }
            #endregion
            #region 스케줄러 관련
            if (Properties.Settings.Default.작업공간정리스케줄러 || Properties.Settings.Default.로그정리스케줄러)
            {
                string SchedulerInfo = "";
                if (Properties.Settings.Default.작업공간정리스케줄러)
                    SchedulerInfo += string.Format("작업공간 정리 스케줄러: {0}일   ", Properties.Settings.Default.작업공간정리주기_일);
                if (Properties.Settings.Default.로그정리스케줄러)
                    SchedulerInfo += string.Format("로그 정리 스케줄러: {0}일", Properties.Settings.Default.로그정리주기_일);
                DevLog.Write("[Scheduler] 스케줄러가 실행중입니다. " + SchedulerInfo, LOG_LEVEL.INFO);

                tScheduler = new System.Windows.Forms.Timer();
                tScheduler.Interval = CalculateTimerInterval(CHECK_INTERVAL);
                tScheduler.Tick += new EventHandler(tScheduler_Tick);
                tScheduler.Start();
            }
            #endregion
            if (Properties.Settings.Default.오피스디버깅모드)
                this.Visible = true;
            else
                this.Visible = false;
            #region Create SocketServer
            socketServer.InitConfig();
            socketServer.CreateServer();

            bool IsResult = false;

            if (checkLicense["HWID"].ToString().Equals(new LicenseInfo().getHWID()))
                IsResult = socketServer.Start();

            if (IsResult)
            {
                DevLog.Write("[Socket] Server Listening...", LOG_LEVEL.INFO);
                DevLog.Write(string.Format("[Socket][INFO] IP: {0}   포트: {1}   프로토콜: {2}   서버이름: {3}", Properties.Settings.Default.바인딩_IP, socketServer.Config.Port, socketServer.Config.Mode, socketServer.Config.Name), LOG_LEVEL.INFO);

                pictureBox1.Image = DocConvert_Server.Properties.Resources.success_icon;
            }
            else
            {
                DevLog.Write(string.Format("[Socket][ERROR] 서버 네트워크 시작 실패, 설정한 IP주소가 일치하지거나 바인딩한 포트가 이미 사용중일경우 발생하는 오류."), LOG_LEVEL.ERROR);
                pictureBox1.Image = DocConvert_Server.Properties.Resources.error_icon;
                return;
            }
            #endregion
            #region Create WebSocketServer
            try
            {
                CancellationTokenSource cancellation = new CancellationTokenSource();
                //var endpoint = new IPEndPoint(IPAddress.Any, 1818);
                var endpoint = new IPEndPoint(IPAddress.Parse(Properties.Settings.Default.바인딩_IP), Properties.Settings.Default.웹소켓포트);
                var options = new WebSocketListenerOptions()
                {
                    WebSocketReceiveTimeout = new TimeSpan(0, 3, 0), // 클라이언트가 서버로 요청했을때 서버가 바쁘면 Timeout
                    WebSocketSendTimeout = new TimeSpan(0, 0, 5), // 클라이언트가 연결을 끊었을때 Timeout
                    NegotiationTimeout = new TimeSpan(0, 3, 0),
                    PingTimeout = new TimeSpan(0, 3, 0),
                    PingMode = PingModes.LatencyControl
                };
                webSocketServer = new WebSocketListener(endpoint, options);
                var rfc6455 = new vtortola.WebSockets.Rfc6455.WebSocketFactoryRfc6455(webSocketServer);
                rfc6455.MessageExtensions.RegisterExtension(new WebSocketDeflateExtension());
                webSocketServer.Standards.RegisterStandard(rfc6455);
                if (checkLicense["HWID"].ToString().Equals(new LicenseInfo().getHWID()))
                    webSocketServer.Start();

                DevLog.Write("[Web Socket] Server Listening...", LOG_LEVEL.INFO);
                DevLog.Write(string.Format("[Web Socket][INFO] IP: {0}   포트: {1}", endpoint.Address, endpoint.Port), LOG_LEVEL.INFO);

                var task = Task.Run(() => AcceptWebSocketClientsAsync(webSocketServer, cancellation.Token));
                pictureBox3.Image = Properties.Resources.success_icon;
            }
            catch (Exception e1)
            {
                logger.Info("변환중 오류발생 자세한 내용은 오류로그 참고");
                logger.Error("==================== Method: " + MethodBase.GetCurrentMethod().Name + " ====================");
                logger.Error(new StackTrace(e1, true));
                logger.Error("변환 실패: " + e1.Message);
                logger.Error("==================== End ====================");
                DevLog.Write("[WebSocketServer][Error] " + e1.Message, LOG_LEVEL.ERROR);
                pictureBox3.Image = Properties.Resources.error_icon;
            }

            // 소켓서버 소멸
            /*Console.ReadKey(true);
            Console.WriteLine("Server stoping");
            cancellation.Cancel();
            task.Wait();*/
            #endregion
            #region Thread
            workProcessTimer.Tick += new EventHandler(OnProcessTimedEvent);
            workProcessTimer.Interval = new TimeSpan(0, 0, 0, 0, 32);
            workProcessTimer.Start();
            new Thread(delegate ()
            {
                while (!this.IsDisposed)
                {
                    toolStripStatusLabel2.Text = string.Format("Socket Session Count: {0}/{1}", socketServer.SessionCount, Properties.Settings.Default.소켓최대세션);
                    toolStripStatusLabel3.Text = string.Format("Web Socket Session Count: {0}", wsSessionCount);

                    if (IsTcpPortAvailable(Properties.Settings.Default.웹소켓포트))
                        pictureBox3.Image = Properties.Resources.success_icon;
                    else
                        pictureBox3.Image = Properties.Resources.error_icon;

                    if (IsTcpPortAvailable(Properties.Settings.Default.파일서버포트))
                        pictureBox2.Image = Properties.Resources.success_icon;
                    else
                        pictureBox2.Image = Properties.Resources.error_icon;
                    Thread.Sleep(1000);
                }
            }).Start();
            #endregion
        }

        #region WebSocket Method
        //웹 소켓으로 접속하는 클라이언트를 받아들입니다.
        async Task AcceptWebSocketClientsAsync(WebSocketListener server, CancellationToken token)
        {
            while (!token.IsCancellationRequested)
            {
                try
                {
                    var ws = await server.AcceptWebSocketAsync(token).ConfigureAwait(false);

                    DevLog.Write("[WebSocket] 클라이언트로 부터 연결요청이 들어옴: " + ws.RemoteEndpoint, LOG_LEVEL.INFO);

                    //소켓이 null 이 아니면, 핸들러를 스타트 합니다.(또 다른 친구가 들어올 수도 있으니 비동기로...)
                    if (ws != null)
                        // await Task.Run(() => HandleConnectionAsync(ws, token)); <== await 시 요청을 받으면 이전 요청이 종료할때까지 대기했다가 실행
                        _ = Task.Run(() => HandleConnectionAsync(ws, token));
                }
                catch (Exception)
                {
                    /*DevLog.Write("[WebSocket] Error Accepting clients: " + aex.GetBaseException().Message, LOG_LEVEL.ERROR);*/
                }
            }
            DevLog.Write("[WebSocket] Server Stop accepting clients", LOG_LEVEL.ERROR);
        }

        async Task HandleConnectionAsync(WebSocket ws, CancellationToken cancellation)
        {
            try
            {
                //연결이 끊기지 않았고, 캔슬이 들어오지 않는 한 루프를 돔.
                while (ws.IsConnected && !cancellation.IsCancellationRequested)
                {
                    ++wsSessionCount;
                    //클라이언트로부터 메시지가 왔는지 비동기읽음
                    string requestInfo = await ws.ReadStringAsync(cancellation).ConfigureAwait(false);

                    if (requestInfo != null)
                    {
                        JObject responseMsg = new JObject();
                        try
                        {
                            DateTime timeTaken = DateTime.Now;
                            JObject requestMsg = JObject.Parse(requestInfo);
                            DevLog.Write(string.Format("\r\n[WebSocket][Client => Server]\r\n{0}\r\n", requestMsg), LOG_LEVEL.INFO);

                            // 문서변환 메소드
                            responseMsg = new Document_Convert().document_Convert(requestInfo);

                            DevLog.Write(string.Format("\r\n[WebSocket][Server => Client]\r\n{0}\r\n", responseMsg), LOG_LEVEL.INFO);

                            TimeSpan curTime = DateTime.Now - timeTaken;
                            DevLog.Write(string.Format("[WebSocket] 작업 소요시간: {0}", curTime.ToString()), LOG_LEVEL.INFO);

                            ws.WriteString(responseMsg.ToString());

                            ws.Close();

                            DevLog.Write("[WebSocket]서버와 연결이 해제되었습니다.", LOG_LEVEL.INFO);
                        }
                        catch (Exception e1)
                        {
                            responseMsg["Msg"] = e1.Message;
                        }
                    }
                    --wsSessionCount;
                }
            }
            catch (Exception e1)
            {
                logger.Info("변환중 오류발생 자세한 내용은 오류로그 참고");
                logger.Error("==================== Method: " + MethodBase.GetCurrentMethod().Name + " ====================");
                logger.Error(new StackTrace(e1, true));
                logger.Error("변환 실패: " + e1.Message);
                logger.Error("==================== End ====================");
                DevLog.Write("[WebSocket] Error Handling connection: " + e1.GetBaseException().Message, LOG_LEVEL.ERROR);
                try { ws.Close(); }
                catch { }
            }
            finally
            {
                //소켓은 Dispose 해 줍니다.
                ws.Dispose();
            }
        }
        #endregion

        #region Logging Method

        private void OnProcessTimedEvent(object sender, EventArgs e)
        {
            int logWorkCount = 0;

            while (true)
            {
                string msg;

                if (DevLog.GetLog(out msg))
                {
                    ++logWorkCount;

                    if (listBoxLog.Items.Count >= Properties.Settings.Default.리스트박스최대로그개수)
                    {
                        listBoxLog.Items.RemoveAt(0);
                    }

                    listBoxLog.Items.Add(msg);

                    if (checkBox1.Checked)
                    {
                        listBoxLog.SelectedIndex = listBoxLog.Items.Count - 1;
                        textBox1.AppendText(msg + "\r\n");
                    }

                    toolStripStatusLabel1.Text = string.Format("LogCount: {0}/{1}", listBoxLog.Items.Count, Properties.Settings.Default.리스트박스최대로그개수);
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

        private void listBoxLog_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!checkBox1.Checked)
            {
                textBox1.Text = "";
                foreach (var logList in listBoxLog.SelectedItems)
                {
                    textBox1.AppendText(logList.ToString() + "\r\n");
                }
            }
        }

        private void listBoxLog_Click(object sender, EventArgs e)
        {
            checkBox1.Checked = false;
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
            e.Cancel = true;
            this.Visible = false;
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (var dialog = new AboutForm())
            {
                dialog.ShowDialog(this);
            }
        }

        private void program_Exit(bool forceExit)
        {
            if (!forceExit)
            {
                if (MessageBox.Show(this, "DocConvert 서버를 종료하시겠습니까?", "경고", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK)
                {
                    try
                    {
                        socketServer.Dispose();
                    }
                    catch (Exception) { }
                    try
                    {
                        webSocketServer.Dispose();
                    }
                    catch (Exception) { }
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
                    socketServer.Dispose();
                }
                catch (Exception) { }
                try
                {
                    webSocketServer.Dispose();
                }
                catch (Exception) { }
                this.Dispose();
                Application.ExitThread();
                Environment.Exit(0);
            }
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            program_Exit(false);
        }

        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            program_Exit(false);
        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            using (var dialog = new AboutForm())
            {
                dialog.ShowDialog(this);
            }
        }

        private void notifyIcon1_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                if (this.Visible)
                    this.Visible = false;
                else
                {
                    this.Visible = true;
                    if (this.WindowState == FormWindowState.Minimized)
                        this.WindowState = FormWindowState.Normal;
                    this.Activate();
                }
            }
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            this.Visible = true;
            this.Activate();
        }

        private int CalculateTimerInterval(int minute)
        {
            if (minute <= 0)
                minute = 60;
            DateTime now = DateTime.Now;

            DateTime future = now.AddMinutes((minute - (now.Minute % minute))).AddSeconds(now.Second * -1).AddMilliseconds(now.Millisecond * -1);

            TimeSpan interval = future - now;

            return (int)interval.TotalMilliseconds;
        }

        private void tScheduler_Tick(object sender, EventArgs e)
        {
            tScheduler.Interval = CalculateTimerInterval(CHECK_INTERVAL);
            if (Properties.Settings.Default.작업공간정리스케줄러)
            {
                DevLog.Write("[Scheduler] 작업공간 정리 스케줄러가 실행되었습니다.", LOG_LEVEL.INFO);
                deleteFolder(Properties.Settings.Default.데이터경로 + @"\workspace", Properties.Settings.Default.작업공간정리주기_일);
                deleteFolder(Properties.Settings.Default.데이터경로 + @"\tmp", Properties.Settings.Default.작업공간정리주기_일);
            }
            if (Properties.Settings.Default.로그정리스케줄러)
            {
                DevLog.Write("[Scheduler] 로그 정리 스케줄러가 실행되었습니다.", LOG_LEVEL.INFO);
                deleteFolder(Application.StartupPath + @"\Log", Properties.Settings.Default.로그정리주기_일);
            }
        }

        public static void deleteFolder(string strPath, int DeletionCycle)
        {
            foreach (string Folder in Directory.GetDirectories(strPath))
                deleteFolder(Folder, DeletionCycle); //재귀함수 호출

            foreach (string file in Directory.GetFiles(strPath))
            {
                FileInfo fi = new FileInfo(file);
                if (fi.LastWriteTime <= DateTime.Now.AddDays(-DeletionCycle))
                {
                    fi.Delete();
                    DevLog.Write("[Scheduler] 파일 " + fi.FullName + " 삭제됨.");
                }
                if (!isFiles(fi.DirectoryName))
                {
                    DirectoryInfo di = new DirectoryInfo(fi.DirectoryName);
                    DevLog.Write("[Scheduler] 폴더 " + fi.DirectoryName + " 삭제됨.");
                    di.Delete(true);
                }
            }
        }

        private static bool isFiles(string dir)
        {
            string[] Directories = Directory.GetDirectories(dir);   // Defalut Folder
            {
                string[] Files = Directory.GetFiles(dir);   // File list Search
                if (Files.Length != 0) return true;

                foreach (string nodeDir in Directories)   // Folder list Search
                {
                    isFiles(nodeDir);   // reSearch
                }
            }
            return false;
        }
    }
}
