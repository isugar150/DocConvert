using NLog;
using System;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace DocConvert_Server
{
    public partial class Form1 : Form
    {
        private static Logger logger = LogManager.GetLogger("DocConvert_Server_Log");
        private byte[] data = new byte[4096];
        private int size = 4096;
        private Socket server;
        Thread t1 = null;
        public Form1()
        {
            InitializeComponent();
            CheckForIllegalCrossThreadCalls = false;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            t1 = new Thread(createSocketServer);
            t1.Start();
        }

        #region 통신관련
        private void createSocketServer()
        {
            try
            {
                server = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
                IPEndPoint iep = new IPEndPoint(IPAddress.Any, 7000);
                server.Bind(iep);
                server.Listen(10);
                server.BeginAccept(new AsyncCallback(AcceptConn), server);
                consoleAppend("Socket Server Started..");
            }
            catch (Exception e1)
            {
                consoleAppend(e1.Message);
            }
        }

        private void AcceptConn(IAsyncResult iar)
        {
            Socket oldserver = (Socket)iar.AsyncState;
            Socket client = oldserver.EndAccept(iar);
            consoleAppend(client.RemoteEndPoint.ToString() + "의 연결 요청 수락");
            // 클라이언트로 문자열 보내는 부분

            string welcome = "서버 메세지: Welcome to my server";
            byte[] message1 = Encoding.UTF8.GetBytes(welcome);
            client.BeginSend(message1, 0, message1.Length, SocketFlags.None, new AsyncCallback(SendData), client);
        }

        private void SendData(IAsyncResult iar)
        {
            try
            {
                Socket client = (Socket)iar.AsyncState;
                int sent = client.EndSend(iar);
                client.BeginReceive(data, 0, size, SocketFlags.None, new AsyncCallback(ReceiveData), client);
            }
            catch (Exception)
            {
                throw;
            }
        }

        private void ReceiveData(IAsyncResult iar)
        {
            //while (!this.IsDisposed)
            //{
                try
                {
                    Socket client = (Socket)iar.AsyncState;
                    int recv = client.EndReceive(iar);

                    if (recv == 0)
                    {
                        client.Close();
                        consoleAppend("연결요청 대기중...");
                        server.BeginAccept(new AsyncCallback(AcceptConn), server);
                        return;
                    }

                    string recvData = Encoding.UTF8.GetString(data, 0, recv);
                    
                    consoleAppend("수신됨: " + recvData);
                    byte[] message2 = Encoding.UTF8.GetBytes(recvData);
                    client.BeginSend(message2, 0, message2.Length, SocketFlags.None, new AsyncCallback(SendData), client);
                }
                catch (Exception e1)
                {
                    consoleAppend(e1.Message);
                    //continue;
                }
            //}
        }
        #endregion

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }

        private void consoleAppend(string text)
        {
            textBox1.AppendText(System.DateTime.Now.ToString("yyyy/MM/dd hh:mm:ss.fff") + "   " + text + "\r\n");
        }
    }
}
