using Newtonsoft.Json.Linq;
using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.ServiceModel.Channels;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using vtortola.WebSockets;

using DocConvert_Server.License;
using NLog;
using System.Reflection;

namespace DocConvert_Server
{
    public partial class Form1 : Form
    {
        private System.Windows.Threading.DispatcherTimer workProcessTimer = new System.Windows.Threading.DispatcherTimer();
        private MainServer socketServer = new MainServer();
        private int wsSessionCount = 0;
        WebSocketListener webSocketServer = null;
        JObject checkLicense;

        private static Logger logger = LogManager.GetLogger("DocConvert_Server_Log");
        public Form1()
        {
            InitializeComponent();
            CheckForIllegalCrossThreadCalls = false;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.Text += " - " + Properties.Settings.Default.serverName;
            textBox1.Text = ("┏━━━┓╋╋╋╋╋┏━━━┓╋╋╋╋╋╋╋╋╋╋╋╋╋┏┓╋┏━━━┓\r\n" +
                             "┗┓┏┓┃╋╋╋╋╋┃┏━┓┃╋╋╋╋╋╋╋╋╋╋╋╋┏┛┗┓┃┏━┓┃\r\n" +
                             "╋┃┃┃┣━━┳━━┫┃╋┗╋━━┳━┓┏┓┏┳━━┳┻┓┏┛┃┗━━┳━━┳━┳┓┏┳━━┳━┓\r\n" +
                             "╋┃┃┃┃┏┓┃┏━┫┃╋┏┫┏┓┃┏┓┫┗┛┃┃━┫┏┫┃╋┗━━┓┃┃━┫┏┫┗┛┃┃━┫┏┛\r\n" +
                             "┏┛┗┛┃┗┛┃┗━┫┗━┛┃┗┛┃┃┃┣┓┏┫┃━┫┃┃┗┓┃┗━┛┃┃━┫┃┗┓┏┫┃━┫┃\r\n" +
                             "┗━━━┻━━┻━━┻━━━┻━━┻┛┗┛┗┛┗━━┻┛┗━┛┗━━━┻━━┻┛╋┗┛┗━━┻┛\r\n");
            #region checkLicense
            try
            {
                Debug.WriteLine(Properties.Settings.Default.LicenseKEY);
                string licenseInfo = LicenseInfo.decryptAES256(Properties.Settings.Default.LicenseKEY, "JmDoCOnVerTerServErJmCoRp");
                checkLicense = JObject.Parse(licenseInfo);
                if (!checkLicense["HWID"].ToString().Equals(new LicenseInfo().getHWID())) { new MessageDialog("라이센스 오류", "라이센스 확인 후 다시시도하세요.", "HWID: " + new LicenseInfo().getHWID()).ShowDialog(this); Application.Exit(); return; }
                if (DateTime.Parse(checkLicense["EndDate"].ToString()) < DateTime.Now) { new MessageDialog("라이센스 오류", "라이센스 날짜가 만료되었습니다. 갱신후 다시시도해주세요.", "HWID: " + new LicenseInfo().getHWID()).ShowDialog(this); Application.Exit(); return; }

                DevLog.Write("나의 하드웨어 ID: " + new LicenseInfo().getHWID(), LOG_LEVEL.INFO);
                DevLog.Write(string.Format("라이센스 만료날짜: {0}", checkLicense["EndDate"].ToString()), LOG_LEVEL.INFO);
            }
            catch (Exception) { new MessageDialog("라이센스 오류", "라이센스키 파싱오류.", "HWID: " + new LicenseInfo().getHWID()).ShowDialog(this); Application.Exit(); return; }
            #endregion
            toolStripStatusLabel4.Text = "IP Address: " + Properties.Settings.Default.serverIP;
            toolStripStatusLabel5.Text = "Socket Port: : " + Properties.Settings.Default.socketPORT.ToString();
            toolStripStatusLabel6.Text = "WebSocket Port: " + Properties.Settings.Default.webSocketPORT.ToString();
            toolStripStatusLabel7.Text = "File Server Port: " + Properties.Settings.Default.fileServerPORT.ToString();
            checkBox1.Checked = Properties.Settings.Default.FollowTail;
            #region Create SocketServer
            socketServer.InitConfig();
            socketServer.CreateServer();
            bool IsResult = false;
            if (checkLicense["HWID"].ToString().Equals(new LicenseInfo().getHWID()))
                IsResult = socketServer.Start();

            if (IsResult)
            {
                DevLog.Write("[Socket Server Listening...]", LOG_LEVEL.INFO);
                DevLog.Write(string.Format("[Socket][INFO] IP: {0}   포트: {1}   프로토콜: {2}   서버이름: {3}", Properties.Settings.Default.serverIP, socketServer.Config.Port, socketServer.Config.Mode, socketServer.Config.Name), LOG_LEVEL.INFO);

                pictureBox1.Image = DocConvert_Server.Properties.Resources.success_icon;
            }
            else
            {
                DevLog.Write(string.Format("[Socket][ERROR] 서버 네트워크 시작 실패"), LOG_LEVEL.ERROR);
                pictureBox1.Image = DocConvert_Server.Properties.Resources.error_icon;
                return;
            }
            #endregion
            #region LogManger
            workProcessTimer.Tick += new EventHandler(OnProcessTimedEvent);
            workProcessTimer.Interval = new TimeSpan(0, 0, 0, 0, 32);
            workProcessTimer.Start();
            new Thread(delegate ()
            {
                while (!this.IsDisposed)
                {
                    toolStripStatusLabel2.Text = string.Format("Socket Session Count: {0}/{1}", socketServer.SessionCount, Properties.Settings.Default.socketSessionCount);
                    toolStripStatusLabel3.Text = string.Format("Web Socket Session Count: {0}/{1}", wsSessionCount, "1");

                    if (IsTcpPortAvailable(Properties.Settings.Default.webSocketPORT))
                        pictureBox3.Image = Properties.Resources.success_icon;
                    else
                    {
                        pictureBox3.Image = Properties.Resources.error_icon;
                    }

                    if (IsTcpPortAvailable(Properties.Settings.Default.fileServerPORT))
                        pictureBox2.Image = Properties.Resources.success_icon;
                    else
                    {
                        pictureBox2.Image = Properties.Resources.error_icon;
                    }
                    Thread.Sleep(1000);
                }
            }).Start();
            #endregion
            #region Create WebSocketServer
            try
            {
                CancellationTokenSource cancellation = new CancellationTokenSource();
                //var endpoint = new IPEndPoint(IPAddress.Any, 1818);
                var endpoint = new IPEndPoint(IPAddress.Parse(Properties.Settings.Default.serverIP), Properties.Settings.Default.webSocketPORT);
                var options = new WebSocketListenerOptions()
                {
                    WebSocketReceiveTimeout = new TimeSpan(0, 3, 0),
                    WebSocketSendTimeout = new TimeSpan(0, 0, 5), // 클라이언트가 연결을 끊었을때 타임아웃
                    NegotiationTimeout = new TimeSpan(0, 3, 0),
                    PingTimeout = new TimeSpan(0, 3, 0),
                    PingMode = PingModes.LatencyControl
                };
                webSocketServer = new WebSocketListener(endpoint, options);
                var rfc6455 = new vtortola.WebSockets.Rfc6455.WebSocketFactoryRfc6455(webSocketServer);
                webSocketServer.Standards.RegisterStandard(rfc6455);
                if (checkLicense["HWID"].ToString().Equals(new LicenseInfo().getHWID()))
                    webSocketServer.Start();

                DevLog.Write("[Web Socket Server Listening...]", LOG_LEVEL.INFO);
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
                        await Task.Run(() => HandleConnectionAsync(ws, token));
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
                            DevLog.Write(string.Format("[Socket] 작업 소요시간: {0}", curTime.ToString()), LOG_LEVEL.INFO);

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

                    if (listBoxLog.Items.Count >= Properties.Settings.Default.LogMaxCount)
                    {
                        listBoxLog.Items.RemoveAt(0);
                    }

                    listBoxLog.Items.Add(msg);

                    if (checkBox1.Checked)
                    {
                        listBoxLog.SelectedIndex = listBoxLog.Items.Count - 1;
                        textBox1.AppendText(msg + "\r\n");
                    }

                    toolStripStatusLabel1.Text = string.Format("LogCount: {0}/{1}", listBoxLog.Items.Count, Properties.Settings.Default.LogMaxCount);
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
            Application.Exit();
        }
    }
}
