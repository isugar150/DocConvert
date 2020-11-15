using DocConvert_Core.IniLib;
using DocConvert_Core.interfaces;
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

namespace DocConvert_Server
{
    public partial class Form1 : Form
    {
        #region 전역변수
        private static Logger logger = LogManager.GetLogger("DocConvert_Server_Log");
        private System.Windows.Threading.DispatcherTimer workProcessTimer = new System.Windows.Threading.DispatcherTimer();
        private System.Windows.Forms.Timer tScheduler;
        private WebSocketListener webSocketServer = null;
        private JObject checkLicense = new JObject();
        private int wsSessionCount = 0;
        private int selectLastIndex = 0;
        private bool checkedFollowTail = false;
        private bool noLicense = false;
        #endregion

        #region 공유변수
        public static bool isHwpConverting = false;
        public static iniProperties IniProperties = new iniProperties();
        #endregion

        public Form1(string[] args)
        {
            #region Init Form
            InitializeComponent();
            CheckForIllegalCrossThreadCalls = false;
            #endregion

            #region Welcome Message
            DevLog.Write2("┏━━━┓╋╋╋╋╋┏━━━┓╋╋╋╋╋╋╋╋╋╋╋╋╋┏┓╋┏━━━┓\r\n" +
                          "┗┓┏┓┃╋╋╋╋╋┃┏━┓┃╋╋╋╋╋╋╋╋╋╋╋╋┏┛┗┓┃┏━┓┃\r\n" +
                          "╋┃┃┃┣━━┳━━┫┃╋┗╋━━┳━┓┏┓┏┳━━┳┻┓┏┛┃┗━━┳━━┳━┳┓┏┳━━┳━┓\r\n" +
                          "╋┃┃┃┃┏┓┃┏━┫┃╋┏┫┏┓┃┏┓┫┗┛┃┃━┫┏┫┃╋┗━━┓┃┃━┫┏┫┗┛┃┃━┫┏┛\r\n" +
                          "┏┛┗┛┃┗┛┃┗━┫┗━┛┃┗┛┃┃┃┣┓┏┫┃━┫┃┃┗┓┃┗━┛┃┃━┫┃┗┓┏┫┃━┫┃\r\n" +
                          "┗━━━┻━━┻━━┻━━━┻━━┻┛┗┛┗┛┗━━┻┛┗━┛┗━━━┻━━┻┛╋┗┛┗━━┻┛", LOG_LEVEL.TRACE);
            #endregion

            #region 아규먼트 파싱
            if (args.Length != 0)
            {
                for(int i = 0; i<args.Length; i++)
                {
                    if (args[i].ToLower().Equals("/minimized"))
                    {
                        this.WindowState = FormWindowState.Minimized;
                    }
                    /*if (args[i].ToLower().Contains("/nolicense"))
                    {
                        if (args[i].Split('=')[1].Equals("JmSoftware"))
                        {
                            noLicense = true;
                            DevLog.Write("라이센스없이 프로그램을 실행하였습니다.\r\n=========================>무단으로 사용할 경우 법적 처벌을 받을 수 있습니다.", LOG_LEVEL.DEBUG);
                            this.Text += " - No License Version ";
                        }
                    }*/
                }
            }
            #endregion

            noLicense = true;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            #region Parse INI
            IniFile pairs = new IniFile();
            while (true)
            {
                try
                {
                    if (new FileInfo("./DocConvert_Server.ini").Exists)
                    {
                        pairs.Load("./DocConvert_Server.ini");
                        IniProperties.LicenseKEY = pairs["DC Server"]["LicenseKEY"].ToString2();
                        IniProperties.BindIP = pairs["DC Server"]["BindIP"].ToString2();
                        IniProperties.WebSocketPort = int.Parse(pairs["DC Server"]["WebSocketPort"].ToString2());
                        IniProperties.FileServerPort = int.Parse(pairs["DC Server"]["FileServerPort"].ToString2());
                        IniProperties.DisplayLogCnt = int.Parse(pairs["DC Server"]["DisplayLogCnt"].ToString2());
                        IniProperties.ClientKEY = pairs["DC Server"]["ClientKEY"].ToString2();
                        IniProperties.DataPath = pairs["DC Server"]["DataPath"].ToString2();
                        IniProperties.ResponseTimeout = pairs["DC Server"]["ResponseTimeout"].ToString2();
                        IniProperties.SchedulerTime = pairs["DC Server"]["SchedulerTime"].ToString2();
                        IniProperties.OfficeDebugModeYn = pairs["DC Server"]["OfficeDebugModeYn"].ToString2().Equals("Y");
                        IniProperties.ShowTextBoxYn = pairs["DC Server"]["ShowTextBoxYn"].ToString2().Equals("Y");
                        IniProperties.FollowTailYn = pairs["DC Server"]["FollowTailYn"].ToString2().Equals("Y");
                        IniProperties.CleanWorkspaceSchedulerYn = pairs["DC Server"]["CleanWorkspaceSchedulerYn"].ToString2().Equals("Y");
                        IniProperties.CleanWorkspaceDay = int.Parse(pairs["DC Server"]["CleanWorkspaceDay"].ToString2());
                        IniProperties.CleanLogSchedulerYn = pairs["DC Server"]["CleanLogSchedulerYn"].ToString2().Equals("Y");
                        IniProperties.CleanLogDay = int.Parse(pairs["DC Server"]["CleanLogDay"].ToString2());
                        IniProperties.ChromiumCaptureYn = pairs["DC Server"]["ChromiumCaptureYn"].ToString2().Equals("Y");
                        IniProperties.WebCaptureTimeout = int.Parse(pairs["DC Server"]["WebCaptureTimeout"].ToString2());

                        IniProperties.DRM_useYn = pairs["DRM Setting"]["DRM useYn"].ToString2().Equals("Y");
                        IniProperties.DRM_Path = pairs["DRM Setting"]["DRM Path"].ToString2();
                        IniProperties.DRM_Result = pairs["DRM Setting"]["DRM Result"].ToString2();
                        IniProperties.DRM_Args = pairs["DRM Setting"]["DRM Args"].ToString2();

                        break;
                    }
                    else
                    {
                        Setting.createSetting();
                        MessageBox.Show("설정 파일을 생성하였습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                catch (Exception e1)
                {
                    if (MessageBox.Show(string.Format("다음 위치에서 예외가 발생했습니다.\r\n{0}\r\n기존 파일은 자동으로 백업합니다. 초기화하시겠습니까?", new StackTrace(e1)), "알림", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                    {
                        FileInfo backupFile = new FileInfo("./DocConvert_Server.ini");
                        backupFile.CopyTo(string.Format("DocConvert_Server.ini.{0}.{1}", DateTime.Now.ToString("yyMMdd"), CurrentTimeMillis()));
                        Setting.createSetting();
                        MessageBox.Show("설정 파일을 초기화하였습니다.\r\n환경 설정을 마무리하세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        Process.Start("notepad", Application.StartupPath + @"\DocConvert_Server.ini");
                        //new MessageDialog("알림", "기존 라이센스번호를 입력하거나 다음 HWID와 함께 판매처로 문의하세요.", "HWID: " + new LicenseInfo().getHWID()).ShowDialog(this);
                        program_Exit(true);
                    }
                    else
                    {
                        program_Exit(true);
                    }
                }
            }
            #endregion

            #region 가상 머신환경 체크
            /*
            try
            {
                new LicenseInfo().getHWID();
            }
            catch (Exception) {
                MessageBox.Show("해당 프로그램은 가상 환경에서 사용이 불가능합니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                program_Exit(true);
            }
            #endregion*/
            /*#region checkLicense
            try
            {
                Debug.WriteLine(IniProperties.LicenseKEY);
                string licenseInfo = LicenseInfo.decryptAES256(IniProperties.LicenseKEY, "JmDoCOnVerTerServErJmCoRp");
                checkLicense = JObject.Parse(licenseInfo);
                if (!checkLicense["HWID"].ToString().Equals(new LicenseInfo().getHWID()) && !noLicense) { new MessageDialog("라이센스 오류", "라이센스 확인 후 다시시도하세요.", "HWID: " + new LicenseInfo().getHWID()).ShowDialog(this); program_Exit(true); return; }
                if (DateTime.Parse(checkLicense["EndDate"].ToString()) < DateTime.Now && !noLicense) { new MessageDialog("라이센스 오류", "라이센스 날짜가 만료되었습니다. 갱신후 다시시도해주세요.", "HWID: " + new LicenseInfo().getHWID() + "\r\n" + "파싱한 날짜: " + DateTime.Parse(checkLicense["EndDate"].ToString()).ToString("yyyy-MM-dd")).ShowDialog(this); program_Exit(true); return; }

                this.Text += " - " + checkLicense["CompanyName"].ToString();

                DevLog.Write("[INFO] 나의 하드웨어 ID: " + new LicenseInfo().getHWID(), LOG_LEVEL.INFO);
                if (!noLicense)
                {
                    DevLog.Write(string.Format("[INFO] 라이센스 만료날짜: {0}", checkLicense["EndDate"].ToString()), LOG_LEVEL.INFO);
                }
            }
            catch (Exception) { new MessageDialog("라이센스 오류", "라이센스키 파싱오류.", "HWID: " + new LicenseInfo().getHWID()).ShowDialog(this); if (!noLicense) { program_Exit(true); return; } }
            */
            #endregion

            #region 초기 설정 및 파싱 데이터 출력
            DirectoryInfo directoryInfo = new DirectoryInfo(IniProperties.DataPath + @"\tmp");
            if (!directoryInfo.Exists)
            {
                directoryInfo.Create();
            }

            toolStripStatusLabel4.Text = "IP Address: " + IniProperties.BindIP;
            toolStripStatusLabel6.Text = "WebSocket Port: " + IniProperties.WebSocketPort;
            toolStripStatusLabel7.Text = "File Server Port: " + IniProperties.FileServerPort;

            checkBox1.Checked = IniProperties.ShowTextBoxYn;
            checkBox2.Checked = IniProperties.FollowTailYn;

            if (checkBox2.Checked)
            {
                IniProperties.FollowTailYn = true;
                listBoxLog.Enabled = false;
            }
            else
            {
                listBoxLog.ClearSelected();
                IniProperties.FollowTailYn = false;
                listBoxLog.Enabled = true;
            }
            if (!checkBox1.Checked)
                checkBox2.Checked = false;

            DevLog.Write("[INFO] 데이터 경로: " + IniProperties.DataPath, LOG_LEVEL.INFO);
            DevLog.Write("[INFO] 클라이언트 키: " + IniProperties.ClientKEY, LOG_LEVEL.INFO);
            DevLog.Write("[INFO] 오피스 디버깅모드: " + IniProperties.OfficeDebugModeYn, LOG_LEVEL.INFO);
            if (IniProperties.ChromiumCaptureYn)
            {
                DevLog.Write("[INFO] 웹 캡쳐 모드: Chromium Capture", LOG_LEVEL.INFO);
            }
            else
            {
                DevLog.Write("[INFO] 웹 캡쳐 모드: Phantom JS", LOG_LEVEL.INFO);
            }

            DevLog.Write("[INFO] 웹 캡쳐 타임아웃: " + IniProperties.WebCaptureTimeout + "초", LOG_LEVEL.INFO);

            DevLog.Write("[INFO] DRM 사용 여부: " + IniProperties.DRM_useYn);
            DevLog.Write("[INFO] DRM 경로: " + IniProperties.DRM_Path);
            DevLog.Write("[INFO] DRM 성공 코드: " + IniProperties.DRM_Result);
            // DRM 아규먼트 
            string[] DRM_args = IniProperties.DRM_Args.Split(',');
            for (int i = 0; i < DRM_args.Length; i++)
            {
                DevLog.Write("[INFO] DRM 아규먼트[" + i + "]: " + DRM_args[i]);
            }
            #endregion

            #region 한글 DLL 레지스트리 등록
            if (File.Exists(Application.StartupPath + @"\FilePathCheckerModuleExample.dll"))
            {
                RegistryKey regKey = Registry.CurrentUser.CreateSubKey(@"Software\HNC\HwpCtrl\Modules", RegistryKeyPermissionCheck.ReadWriteSubTree);
                DevLog.Write("[INFO] 한글 DLL을 레지스트리에 등록하였습니다.", LOG_LEVEL.INFO);
            }
            #endregion

            #region 스케줄러 호출
            string SchedulerInfo = "";
            if (IniProperties.CleanWorkspaceSchedulerYn)
            {
                SchedulerInfo += string.Format("작업공간 정리 스케줄러: {0}일   ", IniProperties.CleanWorkspaceDay);
            }

            if (IniProperties.CleanLogSchedulerYn)
            {
                SchedulerInfo += string.Format("로그 정리 스케줄러: {0}일", IniProperties.CleanLogDay);
            }

            DevLog.Write("[Scheduler] 스케줄러가 실행중입니다. " + SchedulerInfo, LOG_LEVEL.INFO);

            tScheduler = new System.Windows.Forms.Timer();
            tScheduler.Interval = CalculateTimerInterval();
            tScheduler.Tick += new EventHandler(tScheduler_Tick);
            tScheduler.Start();
            DevLog.Write("[Scheduler] 다음 스케줄러 작동시간: " + (DateTime.Now + TimeSpan.FromMilliseconds(tScheduler.Interval)));
            #endregion

            #region Create WebSocketServer
            try
            {
                CancellationTokenSource cancellation = new CancellationTokenSource();
                //var endpoint = new IPEndPoint(IPAddress.Any, 1818);
                IPEndPoint endpoint;
                if (IniProperties.BindIP.Equals("Any"))
                {
                    endpoint = new IPEndPoint(IPAddress.Any, IniProperties.WebSocketPort);
                }
                else
                {
                    endpoint = new IPEndPoint(IPAddress.Parse(IniProperties.BindIP), IniProperties.WebSocketPort);
                }


                string[] timeStr = IniProperties.ResponseTimeout.Split(',');
                int[] time = new int[2];
                time[0] = int.Parse(timeStr[0]); // 분
                time[1] = int.Parse(timeStr[1]); // 초

                DevLog.Write("[Web Socket] 웹 소켓 타임아웃: " + time[0] + "분 " + time[1] + "초", LOG_LEVEL.INFO);

                WebSocketListenerOptions options = new WebSocketListenerOptions()
                {
                    WebSocketReceiveTimeout = new TimeSpan(0, time[0], time[1]), // 클라이언트가 서버로 요청했을때 서버가 바쁘면 Timeout
                    WebSocketSendTimeout = new TimeSpan(0, 1, 0), // 클라이언트가 연결을 끊었을때 Timeout
                    NegotiationTimeout = new TimeSpan(0, time[0], time[1]),
                    PingTimeout = new TimeSpan(0, time[0], time[1]),
                    PingMode = PingModes.LatencyControl
                };
                webSocketServer = new WebSocketListener(endpoint, options);
                vtortola.WebSockets.Rfc6455.WebSocketFactoryRfc6455 rfc6455 = new vtortola.WebSockets.Rfc6455.WebSocketFactoryRfc6455(webSocketServer);
                rfc6455.MessageExtensions.RegisterExtension(new WebSocketDeflateExtension());
                webSocketServer.Standards.RegisterStandard(rfc6455);
                webSocketServer.Start();

                DevLog.Write("[Web Socket] Server Listening...", LOG_LEVEL.INFO);
                DevLog.Write(string.Format("[Web Socket][INFO] IP: {0}   포트: {1}", endpoint.Address, endpoint.Port), LOG_LEVEL.INFO);

                Task task = Task.Run(() => AcceptWebSocketClientsAsync(webSocketServer, cancellation.Token));
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
            #endregion

            #region Thread
            workProcessTimer.Tick += new EventHandler(OnProcessTimedEvent);
            workProcessTimer.Interval = new TimeSpan(0, 0, 0, 0, 32);
            workProcessTimer.Start();
            new Thread(delegate ()
            {
                while (!this.IsDisposed)
                {
                    toolStripStatusLabel3.Text = string.Format("Web Socket Session Count: {0}", wsSessionCount);

                    if (IsTcpPortAvailable(IniProperties.WebSocketPort))
                    {
                        pictureBox3.Image = Properties.Resources.success_icon;
                    }
                    else
                    {
                        pictureBox3.Image = Properties.Resources.error_icon;
                        try
                        {
                            webSocketServer.Start();
                            DevLog.Write("웹 소켓 서버가 비정상적으로 종료되어 재시작 하였습니다.");
                        }
                        catch (Exception e1) { DevLog.Write(e1.Message); }
                    }

                    // 라이센스 체크로직
                    if (!noLicense)
                    {
                        if (DateTime.Parse(checkLicense["EndDate"].ToString()) < DateTime.Now) { webSocketServer.Stop(); new MessageDialog("라이센스 오류", "라이센스 날짜가 만료되어 소켓 LISTEN을 중지하였습니다. 갱신 후 다시시도하세요.", "HWID: " + new LicenseInfo().getHWID() + "\r\n" + "파싱한 날짜: " + DateTime.Parse(checkLicense["EndDate"].ToString()).ToString("yyyy-MM-dd")).ShowDialog(this); program_Exit(true); return; }
                    }
                    
                    Thread.Sleep(1000);
                }
            }).Start();
            #endregion
        }

        #region WebSocket Method
        /// <summary>
        /// 클라이언트가 웹 소켓으로 접속했을때 작동하는 로직
        /// </summary>
        /// <param name="server">웹 소켓 리스너</param>
        /// <param name="token">웹 소켓 토큰</param>
        /// <returns></returns>
        private async Task AcceptWebSocketClientsAsync(WebSocketListener server, CancellationToken token)
        {
            while (!token.IsCancellationRequested)
            {
                try
                {
                    WebSocket ws = await server.AcceptWebSocketAsync(token).ConfigureAwait(false);

                    DevLog.Write(string.Format("[WebSocket] 접속 IP: {0}", ws.RemoteEndpoint.Address), LOG_LEVEL.INFO);

                    // 웹 소켓이 null 이 아니면, 핸들러를 스타트 합니다.(또 다른 친구가 들어올 수도 있으니 비동기로...)
                    if (ws != null)
                    {
                        // await Task.Run(() => HandleConnectionAsync(ws, token)); <== await 시 요청을 받으면 이전 요청이 종료할때까지 대기했다가 실행
                        _ = Task.Run(() => HandleConnectionAsync(ws, token));
                    }
                }
                catch (Exception)
                {
                    /*DevLog.Write("[WebSocket] Error Accepting clients: " + aex.GetBaseException().Message, LOG_LEVEL.ERROR);*/
                }
            }
            DevLog.Write("[WebSocket] Server Stop accepting clients", LOG_LEVEL.ERROR);
        }

        /// <summary>
        /// 클라이언트가 메시지 던졌을때 로직
        /// </summary>
        /// <param name="ws">웹 소켓</param>
        /// <param name="cancellation">토큰</param>
        /// <returns></returns>
        private async Task HandleConnectionAsync(WebSocket ws, CancellationToken cancellation)
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
                            //DevLog.Write(string.Format("\r\n[WebSocket][Client => Server]\r\n{0}\r\n", requestMsg), LOG_LEVEL.INFO);

                            // 문서변환 메소드
                            responseMsg = new Document_Convert().document_Convert(requestInfo);

                            //DevLog.Write(string.Format("\r\n[WebSocket][Server => Client]\r\n{0}\r\n", responseMsg), LOG_LEVEL.INFO);

                            TimeSpan curTime = DateTime.Now - timeTaken;
                            DevLog.Write(string.Format("[WebSocket] 작업 소요시간: {0}", curTime.ToString()), LOG_LEVEL.INFO);

                            ws.WriteString(responseMsg.ToString());

                            ws.Close();

                            DevLog.Write("[WebSocket] 클라이언트와 연결을 해제함.", LOG_LEVEL.INFO);
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
                // 웹 소켓은 Dispose 해 줍니다.
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

                    if (listBoxLog.Items.Count >= IniProperties.DisplayLogCnt)
                    {
                        listBoxLog.Items.RemoveAt(0);
                    }

                    listBoxLog.Items.Add(msg);

                    if (checkBox2.Checked)
                    {
                        listBoxLog.SelectedIndex = listBoxLog.Items.Count - 1;
                    }

                    toolStripStatusLabel1.Text = string.Format("LogCount: {0}/{1}", listBoxLog.Items.Count, IniProperties.DisplayLogCnt);
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
            if (checkBox1.Checked)
            {
                if (!checkedFollowTail)
                {
                    textBox1.Text = "";

                    foreach (object logList in listBoxLog.SelectedItems)
                    {
                        textBox1.AppendText(logList.ToString() + "\r\n");
                    }
                }
                else
                {
                    textBox1.AppendText(listBoxLog.Items[selectLastIndex].ToString() + "\r\n");
                }
            }
        }

        #endregion

        #region Port Method
        /// <summary>
        /// TCP 포트가 살아있는지 확인하는 함수
        /// </summary>
        /// <param name="tcpPort">확인할 TCP포트</param>
        /// <returns>살아잇으면 true</returns>
        public static bool IsTcpPortAvailable(int tcpPort)
        {
            IPGlobalProperties ipgp = IPGlobalProperties.GetIPGlobalProperties();

            // Check LISTENING ports
            IPEndPoint[] endpoints = ipgp.GetActiveTcpListeners();
            foreach (IPEndPoint ep in endpoints)
            {
                if (ep.Port == tcpPort)
                {
                    return true;
                }
            }

            return false;
        }
        #endregion

        #region 스케줄러 메소드

        /// <summary>
        /// 스케줄러 등록시 지정할 타이머
        /// </summary>
        /// <returns>해당 시간까지 Milliseconds</returns>
        private int CalculateTimerInterval()
        {
            string[] timeStr = IniProperties.SchedulerTime.Split(',');
            int[] time = new int[3];
            for (int i = 0; i < 2; i++)
                time[0] = int.Parse(timeStr[0]); // 시
            time[1] = int.Parse(timeStr[1]); // 분
            DateTime timeTaken = new DateTime();
            if (new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, time[0], time[1], 0) < DateTime.Now)
                timeTaken = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, time[0], time[1], 0).AddDays(1);
            else
                timeTaken = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, time[0], time[1], 0);

            TimeSpan curTime = timeTaken - DateTime.Now;

            Debug.WriteLine(timeTaken.ToString("yyyy-MM-dd HH:mm:ss"));
            Debug.WriteLine(curTime);
            return (int)curTime.TotalMilliseconds;
        }

        /// <summary>
        /// 해당 시간이 되면 작업 스케줄러 실행로직
        /// </summary>
        private void tScheduler_Tick(object sender, EventArgs e)
        {
            Thread.Sleep(3000);
            try
            {
                tScheduler.Interval = CalculateTimerInterval();
                DevLog.Write("[Scheduler] 다음 스케줄러 작동시간: " + (DateTime.Now + TimeSpan.FromMilliseconds(tScheduler.Interval)));
            }
            catch (Exception) { tScheduler.Dispose(); DevLog.Write("스케줄러 작동중 오류가 발생하여 비활성화 하였습니다."); }
            if (IniProperties.CleanWorkspaceSchedulerYn)
            {
                DevLog.Write("[Scheduler] 작업공간 정리 스케줄러가 실행되었습니다.", LOG_LEVEL.INFO);
                DirectoryInfo workspace = new DirectoryInfo(IniProperties.DataPath + @"\workspace");
                deleteFolder(workspace.FullName, IniProperties.CleanWorkspaceDay);
                if (!workspace.Exists)
                    workspace.Create();
            }
            DirectoryInfo tmp = new DirectoryInfo(IniProperties.DataPath + @"\tmp");
            deleteFolder(tmp.FullName, 1);
            if (!tmp.Exists)
                tmp.Create();
            if (IniProperties.CleanLogSchedulerYn)
            {
                DevLog.Write("[Scheduler] 로그 정리 스케줄러가 실행되었습니다.", LOG_LEVEL.INFO);
                deleteFolder(Application.StartupPath + @"\Log", IniProperties.CleanLogDay);
            }
        }

        /// <summary>
        /// 수정날짜가 설정한 일수가 지났으면 디렉토리와 안의 파일까지 삭제하는 로직
        /// </summary>
        /// <param name="strPath">대상 경로</param>
        /// <param name="DeletionCycle">지난 일수</param>
        public static void deleteFolder(string strPath, int DeletionCycle)
        {
            foreach (string Folder in Directory.GetDirectories(strPath))
            {
                deleteFolder(Folder, DeletionCycle); //재귀함수 호출
            }

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

        /// <summary>
        /// 파일인지 디렉토리인지 확인하는 로직.
        /// </summary>
        /// <param name="dir"></param>
        /// <returns></returns>
        private static bool isFiles(string dir)
        {
            string[] Directories = Directory.GetDirectories(dir);   // Defalut Folder
            {
                string[] Files = Directory.GetFiles(dir);   // File list Search
                if (Files.Length != 0)
                {
                    return true;
                }

                foreach (string nodeDir in Directories)   // Folder list Search
                {
                    isFiles(nodeDir);   // reSearch
                }
            }
            return false;
        }
        #endregion

        #region 컴포넌트 이벤트

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (AboutForm dialog = new AboutForm())
            {
                dialog.ShowDialog(this);
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
            program_Exit(false);
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            program_Exit(false);
        }

        private void toolStripMenuItem4_Click(object sender, EventArgs e)
        {
            if (this.TopMost)
            {
                this.TopMost = false;
            }
            else
            {
                this.TopMost = true;
            }
        }

        /// <summary>
        /// 프로그램 종료시 로직 Application.Exit(0) 안먹음
        /// </summary>
        /// <param name="forceExit">다이얼로그 띄우고 종료할건지 강제 종료할건지 여부</param>
        private void program_Exit(bool forceExit)
        {
            if (!forceExit)
            {
                if (MessageBox.Show(this, "DocConvert 서버를 종료하시겠습니까?", "경고", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK)
                {
                    if (webSocketServer != null)
                    {
                        DevLog.Write("[INFO] 웹 소켓 서버를 종료합니다.", LOG_LEVEL.INFO);
                        webSocketServer.Stop();
                        webSocketServer.Dispose();
                    }

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
                if (webSocketServer != null)
                {
                    DevLog.Write("[INFO] 웹 소켓 서버를 종료합니다.", LOG_LEVEL.INFO);
                    webSocketServer.Stop();
                    webSocketServer.Dispose();
                }

                try
                {
                    Application.ExitThread();
                    Environment.Exit(0);
                }
                catch (Exception) { }
            }
        }

        private void settingToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DevLog.Write("설정 파일을 실행하였습니다. 변경 내용을 적용하려면 서버를 재시작해야합니다.", LOG_LEVEL.INFO);
            Process.Start("notepad", Application.StartupPath + @"\DocConvert_Server.ini");
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                splitContainer2.SplitterDistance = 300;
                checkBox2.Checked = IniProperties.FollowTailYn;
                checkBox2.Enabled = true;
            }
            else
            {
                splitContainer2.SplitterDistance = splitContainer2.Width;
                checkBox2.Checked = false;
                checkBox2.Enabled = false;
                textBox1.Text = "";
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                listBoxLog.ClearSelected();
                textBox1.Text = "";
                checkedFollowTail = true;
                for (int i = 0; i < listBoxLog.Items.Count; i++)
                {
                    selectLastIndex = i;
                    listBoxLog.SelectedIndex = i;
                }
                checkedFollowTail = false;
                IniProperties.FollowTailYn = true;
                listBoxLog.Enabled = false;
            }
            else
            {
                listBoxLog.ClearSelected();
                IniProperties.FollowTailYn = false;
                listBoxLog.Enabled = true;
            }
        }

        #endregion

        public static long CurrentTimeMillis()
        { 
            DateTime Jan1st1970 = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
            return (long)(DateTime.UtcNow - Jan1st1970).TotalMilliseconds;
        }
    }
}
