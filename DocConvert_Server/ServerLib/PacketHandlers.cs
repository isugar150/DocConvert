using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

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
            DevLog.Write(string.Format("\r\n[Client => Server]\r\n{0}\r\n", Encoding.Unicode.GetString(requestInfo.Body)), LOG_LEVEL.INFO); // 클라이언트가 서버로 보낸 메시지
            JObject responseValue = new JObject();
            JObject requestValue = new JObject();

            try
            {
                responseValue = JObject.Parse(Encoding.Unicode.GetString(requestInfo.Body)); // 요청받은 JSON 파싱

                if (!responseValue["KEY"].ToString().Equals(Properties.Settings.Default.key))
                {
                    requestValue["URL"] = null;
                    requestValue["isSuccess"] = false;
                    requestValue["msg"] = "키가 유효하지 않습니다.";
                    return;
                }

                requestValue["URL"] = Properties.Settings.Default.serverIP + "/";
                requestValue["isSuccess"] = true;
                requestValue["msg"] = "Success";
            } catch(Exception e1)
            {
                requestValue["URL"] = null;
                requestValue["isSuccess"] = false;
                requestValue["msg"] = e1.Message;
            }

            DevLog.Write(string.Format("\r\n[Server => Client]\r\n{0}\r\n", requestValue.ToString()), LOG_LEVEL.INFO); // 클라이언트가 서버로 보낸 메시지

            List<byte> dataSource = new List<byte>();
            dataSource.AddRange(BitConverter.GetBytes((int)PACKETID.REQ_ECHO));
            dataSource.AddRange(BitConverter.GetBytes(requestInfo.Body.Length));
            dataSource.AddRange(Encoding.Unicode.GetBytes(requestValue.ToString()));

            session.Send(dataSource.ToArray(), 0, dataSource.Count);
        }
    }

    public class PK_ECHO
    {
        public string msg;
    }
}
