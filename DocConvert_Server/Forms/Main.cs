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
            checkBox1.Checked = Properties.Settings.Default.FollowTail;
            #region Create SocketServer
            socketServer.InitConfig();
            socketServer.CreateServer();
            var IsResult = socketServer.Start();

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
                        toolStripStatusLabel2.Text = string.Format("        Socket Session Count: {0}/{1}", socketServer.SessionCount, Properties.Settings.Default.socketSessionCount);

                        if (IsTcpPortAvailable(Properties.Settings.Default.webSocketPORT))
                            pictureBox3.Image = Properties.Resources.success_icon;
                        else
                            pictureBox3.Image = Properties.Resources.error_icon;

                        if (IsTcpPortAvailable(Properties.Settings.Default.fileServerPORT))
                            pictureBox2.Image = Properties.Resources.success_icon;
                        else
                            pictureBox2.Image = Properties.Resources.error_icon;
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
                    WebSocketSendTimeout = new TimeSpan(0, 3, 0),
                    NegotiationTimeout = new TimeSpan(0, 3, 0),
                    PingTimeout = new TimeSpan(0, 3, 0),
                    PingMode = PingModes.LatencyControl
                };
                WebSocketListener server = new WebSocketListener(endpoint, options);
                var rfc6455 = new vtortola.WebSockets.Rfc6455.WebSocketFactoryRfc6455(server);
                server.Standards.RegisterStandard(rfc6455);
                server.Start();

                DevLog.Write("[Web Socket Server Listening...]", LOG_LEVEL.INFO);
                DevLog.Write(string.Format("[Web Socket][INFO] IP: {0}   포트: {1}", endpoint.Address, endpoint.Port), LOG_LEVEL.INFO);

                var task = Task.Run(() => AcceptWebSocketClientsAsync(server, cancellation.Token));
                pictureBox3.Image = Properties.Resources.success_icon;
            } catch (Exception e1)
            {
                DevLog.Write("[WebSocketServer][Error] " + e1.Message);
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
                    //클라이언트가 들어왔을까요?
                    var ws = await server.AcceptWebSocketAsync(token).ConfigureAwait(false);

                    DevLog.Write("[WebSocket] 클라이언트로 부터 연결요청이 들어옴: " + ws.RemoteEndpoint);

                    //소켓이 null 이 아니면, 핸들러를 스타트 합니다.(또 다른 친구가 들어올 수도 있으니 비동기로...)
                    if (ws != null)
                        await Task.Run(() => HandleConnectionAsync(ws, token));
                }
                catch (Exception aex)
                {
                    DevLog.Write("[WebSocket] Error Accepting clients: " + aex.GetBaseException().Message);
                }
            }
            DevLog.Write("[WebSocket] Server Stop accepting clients");
        }

        async Task HandleConnectionAsync(WebSocket ws, CancellationToken cancellation)
        {
            try
            {
                //연결이 끊기지 않았고, 캔슬이 들어오지 않는 한 루프를 돕니다.
                while (ws.IsConnected && !cancellation.IsCancellationRequested)
                {
                    //클라이언트로부터 메시지가 왔는지 비동기로 읽어요.
                    string requestInfo = await ws.ReadStringAsync(cancellation).ConfigureAwait(false);

                    //읽은 메시지가 null 이 아니면, 뭔가를 처리 합니다. 이 코드는 그냥 ack 라고 에코를 보내도록 만들었습니다.
                    if (requestInfo != null)
                    {
                        JObject responseMsg = new JObject();
                        try
                        {
                            DateTime timeTaken = DateTime.Now;
                            JObject requestMsg = JObject.Parse(requestInfo);
                            DevLog.Write(string.Format("\r\n[WebSocket][Client => Server]\r\n{0}\r\n", requestMsg));

                            // 문서변환 메소드
                            responseMsg = new Document_Convert().document_Convert(requestInfo);

                            DevLog.Write(string.Format("\r\n[WebSocket][Server => Client]\r\n{0}\r\n", responseMsg));

                            ws.WriteString(responseMsg.ToString());

                            ws.Close();

                            DevLog.Write("[WebSocket]서버와 연결이 해제되었습니다.", LOG_LEVEL.INFO);
                        } catch(Exception e1)
                        {
                            responseMsg["Msg"] = e1.Message;
                        }
                    }
                }
            }
            catch (Exception aex)
            {
                DevLog.Write("[WebSocket] Error Handling connection: " + aex.GetBaseException().Message);
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

                    toolStripStatusLabel1.Text = string.Format("LogCount: {0}/{1}", listBoxLog.Items.Count, Properties.Settings.Default.LogMaxCount);

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
            socketServer.Dispose();
            Application.Exit();
        }
    }
}
