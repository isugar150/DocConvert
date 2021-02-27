using DocConvert_Console.Common;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;

namespace DocConvert_Console.Network
{
    public class AsyncSocketServer : Socket
    {
        public AsyncSocketServer(string bindIP, int socketPORT) : base(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
        {
            base.Bind(new IPEndPoint(IPAddress.Parse(bindIP), socketPORT));
            base.Listen(10); // 백로깅
            base.AcceptAsync(new AsyncSocketEvent(this));
        }
    }
    public class AsyncSocketEvent : SocketAsyncEventArgs
    {
        private Socket socket;
        public AsyncSocketEvent(Socket socket)
        {
            this.socket = socket;
            base.UserToken = socket;
            base.Completed += ClientConnected;
        }
        private void ClientConnected(object sender, SocketAsyncEventArgs e)
        {
            // 접속이 완료되면, Client Event를 생성하여 Receive이벤트를 생성한다.
            var client = new Client(e.AcceptSocket);
            // 서버 Event에 cilent를 제거한다.
            e.AcceptSocket = null;
            // Client로부터 Accept이 되면 이벤트를 발생시킨다. (IOCP로 넣는 것)
            this.socket.AcceptAsync(e);
        }
    }

    class Client : SocketAsyncEventArgs
    {
        // 메시지는 개행으로 구분한다.
        private static char CR = (char)0x0D;
        private static char LF = (char)0x0A;
        private Socket socket;
        // 메시지를 모으기 위한 버퍼
        private StringBuilder sb = new StringBuilder();
        private IPEndPoint remoteAddr;
        public Client(Socket socket)
        {
            this.socket = socket;
            // 메모리 버퍼를 초기화 한다. 크기는 1024이다
            base.SetBuffer(new byte[1024], 0, 1024);
            base.UserToken = socket;
            // 메시지가 오면 이벤트를 발생시킨다. (IOCP로 꺼내는 것)
            base.Completed += Client_Completed; ;
            // 메시지가 오면 이벤트를 발생시킨다. (IOCP로 넣는 것)
            this.socket.ReceiveAsync(this);
            // 접속 환영 메시지
            remoteAddr = (IPEndPoint)socket.RemoteEndPoint;
            LogMgr.Write("Connect Client From: " + remoteAddr.Address.ToString() + ":" + remoteAddr.Port.ToString() + "  Connection time: " + DateTime.Now.ToString(), LOG_LEVEL.INFO);
            this.Send("Welcome DocConvert server!\r\n>");
        }


        private void Client_Completed(object sender, SocketAsyncEventArgs e)
        {
            // 접속이 연결되어 있으면...
            if (socket.Connected && base.BytesTransferred > 0)
            {
                // 수신 데이터는 e.Buffer에 있다.
                byte[] data = e.Buffer;
                // 데이터를 string으로 변환한다.
                //string msg = Encoding.ASCII.GetString(data);
                string msg = Encoding.UTF8.GetString(data);
                // 메모리 버퍼를 초기화 한다. 크기는 1024이다
                base.SetBuffer(new byte[1024], 0, 1024);
                // 버퍼의 공백은 없앤다.
                sb.Append(msg.Trim('\0'));
                // 메시지의 끝이 이스케이프 \r\n의 형태이면 서버에 표시한다.
                if (sb.Length >= 2 && sb[sb.Length - 2] == CR && sb[sb.Length - 1] == LF)
                {
                    // 개행은 없애고..
                    sb.Length = sb.Length - 2;
                    msg = sb.ToString();

                    LogMgr.Write("Received messages from " + remoteAddr.Address.ToString() + ":" + remoteAddr.Port.ToString() + "\r\n" + msg, LOG_LEVEL.DEBUG);
                    // Client로 Echo를 보낸다.

                    JObject responseMsg = new JObject();
                    try
                    {
                        DateTime timeTaken = DateTime.Now;

                        // 문서변환 메소드
                        responseMsg = new Flow.Document_Convert().docConvert(msg);


                        TimeSpan curTime = DateTime.Now - timeTaken;
                        LogMgr.Write(string.Format("작업 소요시간: {0}", curTime.ToString()), LOG_LEVEL.INFO);

                        Send(responseMsg.ToString());

                        LogMgr.Write("[WebSocket] 클라이언트와 연결을 해제함.", LOG_LEVEL.INFO);
                    }
                    catch (Exception e1)
                    {
                        responseMsg["Msg"] = e1.Message;
                    }

                    // 만약 메시지가 exit이면 접속을 끊는다.
                    if ("exit".Equals(msg, StringComparison.OrdinalIgnoreCase))
                    {
                        // 접속 종료 메시지
                        LogMgr.Write("Disconnected : (From:" + remoteAddr.Address.ToString() + ":" + remoteAddr.Port.ToString() + "Disconnected time: " + DateTime.Now.ToString() + " Token: " + this.UserToken, LOG_LEVEL.INFO);
                        // 접속을 중단한다.
                        socket.DisconnectAsync(this);
                        return;
                    }
                    // 버퍼를 비운다.
                    sb.Clear();
                }
                // 메시지가 오면 이벤트를 발생시킨다. (IOCP로 넣는 것)
                this.socket.ReceiveAsync(this);
            }
            else
            {
                // 접속이 끊겼다..
                LogMgr.Write("Disconnected : (From:" + remoteAddr.Address.ToString() + ":" + remoteAddr.Port.ToString() + "Disconnected time: " + DateTime.Now.ToString() + " Token: " + this.UserToken, LOG_LEVEL.INFO);
            }
        }

        private void Send(String msg)
        {
            //byte[] sendData = Encoding.ASCII.GetBytes(msg);
            byte[] sendData = Encoding.UTF8.GetBytes(msg);
            socket.Send(sendData, sendData.Length, SocketFlags.None);
        }
    }
}
