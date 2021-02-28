using DocConvert.Common;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NLog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;

namespace DocConvert.Network
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
            this.Send("Welcome DocConvert server!\r\n");
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
                    JObject requestMsg = new JObject();
                    bool isjson = false;

                    try
                    {
                        requestMsg = JObject.Parse(msg);
                        isjson = true;
                    }
                    catch (JsonReaderException)
                    {
                        LogMgr.Write("Received messages from " + remoteAddr.Address.ToString() + ":" + remoteAddr.Port.ToString() + "\r\n" + msg, LOG_LEVEL.DEBUG);

                        Send(msg + "\r\n");

                        // 만약 메시지가 exit이면 접속을 끊는다.
                        if ("exit".Equals(msg, StringComparison.OrdinalIgnoreCase))
                        {
                            // 접속을 중단한다.
                            socket.DisconnectAsync(this);
                            return;
                        }
                    }

                    if (isjson)
                    {
                        JObject responseMsg = new JObject();
                        DateTime timeTaken = DateTime.Now;
                        try
                        {
                            LogMgr.Write("Received messages from " + remoteAddr.Address.ToString() + ":" + remoteAddr.Port.ToString() + "\r\n" + requestMsg.ToString(), LOG_LEVEL.DEBUG);

                            string method = requestMsg["Method"]?.ToString().Trim() ?? ""; // 변환 메소드
                            string clientKey = requestMsg["ClientKey"]?.ToString().Trim() ?? "";  // 신뢰하는 클라이언트 확인.
                            string fileName = requestMsg["FileName"]?.ToString().Trim() ?? "";  // 파일 이름
                            string docPassword = requestMsg["DocPassword"]?.ToString().Trim() ?? "";  // 해당 파일 암호여부
                            int convertImg = int.Parse(requestMsg["ConvertImg"]?.ToString().Trim() ?? "0");  // 이미지 변환 여부 0:변환 안함, 1:jpeg, 2:png, 3:bmp
                            string drmUseYn = requestMsg["DRM_UseYn"]?.ToString().Trim() ?? "";  // DRM 사용 여부
                            string drmType = requestMsg["DRM_Type"]?.ToString().Trim() ?? ""; // DRM 타입 

                            // 클라이언트 키 일치하는지 확인
                            if (!clientKey.Equals(Program.IniProperties.Client_KEY))
                            {
                                responseMsg["ResultCode"] = define.INVALID_CLIENT_KEY_ERROR.ToString();
                                responseMsg["Message"] = "Invalid Client Key";
                                goto sendPoint;
                            }

                            // 문서 변환
                            if (method.Equals("DocConvert"))
                                responseMsg = new Flow.DocConvert()._DocConvert(fileName, docPassword, convertImg, drmUseYn, drmType);
                            else
                            {
                                responseMsg["ResultCode"] = define.INVALID_METHOD_ERROR.ToString();
                                responseMsg["Message"] = "Invalid Method";
                            }

                            sendPoint:

                            LogMgr.Write("Response messages to " + remoteAddr.Address.ToString() + ":" + remoteAddr.Port.ToString() + "\r\n" + responseMsg.ToString(), LOG_LEVEL.DEBUG);

                            TimeSpan curTime = DateTime.Now - timeTaken;
                            LogMgr.Write(string.Format("Processing time: {0}", curTime.ToString(@"ss\.fff")), LOG_LEVEL.DEBUG);

                            Send(responseMsg.ToString());

                            socket.DisconnectAsync(this);
                            return;
                        }
                        catch (Exception e1)
                        {
                            LogMgr.Write("ERROR CODE: " + define.UNDEFINE_ERROR.ToString(), ConsoleColor.Red, LOG_LEVEL.ERROR);
                            LogMgr.Write("ERROR MESSAGE: " + e1.Message, ConsoleColor.Red, LOG_LEVEL.ERROR);
                            if(LogMgr.getLogLevel("DocConvert_Log").Equals("DEBUG"))
                                LogMgr.Write("ERROR STACKTRACE\r\n" + e1.StackTrace, ConsoleColor.Red, LOG_LEVEL.ERROR);

                            TimeSpan curTime = DateTime.Now - timeTaken;
                            LogMgr.Write(string.Format("Processing time: {0}", curTime.ToString(@"ss\.fff")), LOG_LEVEL.DEBUG);

                            responseMsg["ResultCode"] = define.UNDEFINE_ERROR.ToString();
                            responseMsg["Message"] = e1.Message;
                            Send(responseMsg.ToString());

                            try
                            {
                                socket.DisconnectAsync(this);
                                return;
                            } catch(Exception e2)
                            {
                                LogMgr.Write("ERROR CODE: " + define.UNDEFINE_ERROR.ToString(), ConsoleColor.Red, LOG_LEVEL.ERROR);
                                LogMgr.Write("ERROR MESSAGE: " + e2.Message, ConsoleColor.Red, LOG_LEVEL.ERROR);
                                if (LogMgr.getLogLevel("DocConvert_Log").Equals("DEBUG"))
                                    LogMgr.Write("ERROR STACKTRACE\r\n" + e2.StackTrace, ConsoleColor.Red, LOG_LEVEL.ERROR);
                            }
                        }
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
                LogMgr.Write("Disconnected From:" + remoteAddr.Address.ToString() + ":" + remoteAddr.Port.ToString() + " Time: " + DateTime.Now.ToString(), LOG_LEVEL.INFO);
            }
        }

        private void Send(String msg)
        {
            //byte[] sendData = Encoding.ASCII.GetBytes(msg);
            try
            {
                byte[] sendData = Encoding.UTF8.GetBytes(msg);
                socket.Send(sendData, sendData.Length, SocketFlags.None);
            } catch(Exception e1)
            {
                LogMgr.Write("ERROR CODE: " + define.UNDEFINE_ERROR.ToString(), ConsoleColor.Red, LOG_LEVEL.ERROR);
                LogMgr.Write("ERROR MESSAGE: " + e1.Message, ConsoleColor.Red, LOG_LEVEL.ERROR);
                if (LogMgr.getLogLevel("DocConvert_Log").Equals("DEBUG"))
                    LogMgr.Write("ERROR STACKTRACE\r\n" + e1.StackTrace, ConsoleColor.Red, LOG_LEVEL.ERROR);

                try
                {
                    socket.DisconnectAsync(this);
                    return;
                }
                catch (Exception e2)
                {
                    LogMgr.Write("ERROR CODE: " + define.UNDEFINE_ERROR.ToString(), ConsoleColor.Red, LOG_LEVEL.ERROR);
                    LogMgr.Write("ERROR MESSAGE: " + e2.Message, ConsoleColor.Red, LOG_LEVEL.ERROR);
                    if (LogMgr.getLogLevel("DocConvert_Log").Equals("DEBUG"))
                        LogMgr.Write("ERROR STACKTRACE\r\n" + e2.StackTrace, ConsoleColor.Red, LOG_LEVEL.ERROR);
                }
            }
        }
    }
}
