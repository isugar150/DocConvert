using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

using DocConvert_Core.imageLib;
using DocConvert_Core.interfaces;
using DocConvert_Core.OfficeLib;
using DocConvert_Core.HWPLib;

using System.Security.Cryptography;
using System.IO;
using System.Threading;

namespace DocConvert_Server
{
    public class PacketData
    {
        public NetworkSession session;
        public EFBinaryRequestInfo reqInfo;
    }

    public enum PACKETID : int
    {
        REQ_ECHO = 1,
    }
    public class CommonHandler
    {
        public void RequestMsg(NetworkSession session, EFBinaryRequestInfo requestInfo)
        {
            DevLog.Write(string.Format("\r\n[Socket][Client => Server]\r\n{0}\r\n", Encoding.Unicode.GetString(requestInfo.Body)), LOG_LEVEL.INFO); // 클라이언트가 서버로 보낸 메시지

            DateTime timeTaken = DateTime.Now;

            // 문서변환 메소드
            JObject responseMsg = new Document_Convert().document_Convert(Encoding.Unicode.GetString(requestInfo.Body));

            DevLog.Write(string.Format("\r\n[Socket][Server => Client]\r\n{0}\r\n", responseMsg.ToString()), LOG_LEVEL.INFO); // 클라이언트가 서버로 보낸 메시지

            TimeSpan curTime = DateTime.Now - timeTaken;
            DevLog.Write(string.Format("[Socket] 작업 소요시간: {0}", curTime.ToString()), LOG_LEVEL.DEBUG);

            List<byte> dataSource = new List<byte>();
            dataSource.AddRange(BitConverter.GetBytes((int)PACKETID.REQ_ECHO));
            dataSource.AddRange(BitConverter.GetBytes(requestInfo.Body.Length));
            dataSource.AddRange(Encoding.Unicode.GetBytes(responseMsg.ToString()));

            session.Send(dataSource.ToArray(), 0, dataSource.Count);
        }
    }

    public class PK_ECHO
    {
        public string msg;
    }
}
