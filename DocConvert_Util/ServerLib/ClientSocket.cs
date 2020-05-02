using System;
using System.Net;
using System.Net.Sockets;

namespace DocConvert_Util
{
    public class ClientSocket
    {
        public Socket socket = null;    //server 연결을 위한 ClientSock

        public string LatestErrorMsg;
        public string LatestReceiveMsg;

        /// <summary>
        /// 소켓으로 서버에 접속하는 함수
        /// </summary>
        /// <param name="IP">타겟 서버 아이피 주소</param>
        /// <param name="PORT">타겟 서버 포트</param>
        /// <returns>접속 성공여부</returns>
        public bool conn(string IP, int PORT)
        {
            IPAddress serverIP = IPAddress.Parse(IP);
            int serverPort = PORT;

            //Socket 생성(생성 안되면, SocketException 발생!!!!!!)
            socket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);

            //Socket 연결(연결 안되면, SocketException 발생!!!!!!)
            this.socket.Connect(new IPEndPoint(serverIP, serverPort));

            if (socket == null || socket.Connected == false)
            {
                return false;
            }

            return true;
        }

        /// <summary>
        /// 스트림에서 읽어오기
        /// </summary>
        public Tuple<int, byte[]> s_read()
        {
            try
            {
                byte[] getbyte = new byte[4096];
                int nRecv = socket.Receive(getbyte, 0, getbyte.Length, SocketFlags.None);
                return new Tuple<int, byte[]>(nRecv, getbyte);
            }
            catch (SocketException se)
            {
                LatestErrorMsg = se.ToString();
            }

            return null;
        }

        /// <summary>
        /// 스트림에 쓰기
        /// </summary>
        /// <param name="senddata">서버로 보낼 Byte</param>
        public void s_write(byte[] senddata)
        {
            try
            {
                if (socket != null && socket.Connected) //연결상태 유무 확인
                {
                    socket.Send(senddata, 0, senddata.Length, SocketFlags.None);
                }
                else
                {
                    LatestErrorMsg = "소켓 서버에 접속하지못했습니다.";
                }
            }
            catch (SocketException se)
            {
                LatestErrorMsg = se.Message.ToString();
            }
        }

        /// <summary>
        /// 소켓과 스트림 닫기
        /// </summary>
        public void close()
        {
            socket.Close();
        }
    }
}
