﻿using SuperSocket.SocketBase;
using SuperSocket.SocketBase.Config;
using SuperSocket.SocketBase.Logging;
using SuperSocket.SocketBase.Protocol;
using System;
using System.Collections.Generic;

namespace DocConvert_Server
{
    class MainServer : AppServer<NetworkSession, EFBinaryRequestInfo>
    {
        Dictionary<int, Action<NetworkSession, EFBinaryRequestInfo>> HandlerMap = new Dictionary<int, Action<NetworkSession, EFBinaryRequestInfo>>();
        CommonHandler CommonHan = new CommonHandler();

        IServerConfig m_Config;


        public MainServer()
            : base(new DefaultReceiveFilterFactory<ReceiveFilter, EFBinaryRequestInfo>())
        {
            NewSessionConnected += new SessionHandler<NetworkSession>(OnConnected);
            SessionClosed += new SessionHandler<NetworkSession, CloseReason>(OnClosed);
            NewRequestReceived += new RequestHandler<NetworkSession, EFBinaryRequestInfo>(RequestReceived);
        }


        void RegistHandler()
        {
            HandlerMap.Add((int)PACKETID.REQ_ECHO, CommonHan.RequestMsg);

            DevLog.Write(string.Format("[Socket] 핸들러 등록 완료"), LOG_LEVEL.INFO);
        }

        public void InitConfig()
        {
            m_Config = new ServerConfig
            {
                Port = Properties.Settings.Default.소켓서버포트,
                Ip = Properties.Settings.Default.바인딩_IP,
                MaxConnectionNumber = Properties.Settings.Default.소켓최대세션,
                Mode = SocketMode.Tcp,
                Name = Properties.Settings.Default.서버이름
            };
        }

        public void CreateServer()
        {
            bool bResult = Setup(new RootConfig(), m_Config, logFactory: new Log4NetLogFactory());

            if (bResult == false)
            {
                DevLog.Write(string.Format("[Socket][ERROR] 서버 네트워크 설정 실패"), LOG_LEVEL.ERROR);
                return;
            }

            RegistHandler();

            DevLog.Write(string.Format("[Socket] 서버 생성 성공"), LOG_LEVEL.INFO);
        }

        public bool IsRunning(ServerState eCurState)
        {
            if (eCurState == ServerState.Running)
            {
                return true;
            }

            return false;
        }

        void OnConnected(NetworkSession session)
        {
            DevLog.Write(string.Format("[Socket] 접속IP: {0}, 세션ID: {1}", session.RemoteEndPoint.Address, session.SessionID), LOG_LEVEL.DEBUG);
        }

        void OnClosed(NetworkSession session, CloseReason reason)
        {
            DevLog.Write(string.Format("[Socket] 세션ID: {0} 접속해제: {1}", session.SessionID, reason.ToString()), LOG_LEVEL.INFO);
        }

        void RequestReceived(NetworkSession session, EFBinaryRequestInfo reqInfo)
        {
            DevLog.Write(string.Format("[Socket] 세션ID: {0} 받은 데이터 크기: {1}, ThreadId: {2}", session.SessionID, reqInfo.Body.Length, System.Threading.Thread.CurrentThread.ManagedThreadId), LOG_LEVEL.DEBUG);


            var PacketID = reqInfo.PacketID;
            var value1 = reqInfo.Value1;
            var value2 = reqInfo.Value2;

            if (HandlerMap.ContainsKey(PacketID))
            {
                HandlerMap[PacketID](session, reqInfo);
            }
            else
            {
                DevLog.Write(string.Format("[Socket]세션ID: {0} 받은 데이터 크기: {1}", session.SessionID, reqInfo.Body.Length), LOG_LEVEL.DEBUG);
            }
        }
    }


    public class NetworkSession : AppSession<NetworkSession, EFBinaryRequestInfo>
    {
    }
}
