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
            JObject responseMsg = new JObject();

            try
            {
                JObject requestMsg = JObject.Parse(Encoding.Unicode.GetString(requestInfo.Body)); // 요청받은 JSON 파싱

                if (!requestMsg["KEY"].ToString().Equals(Properties.Settings.Default.key))
                {
                    responseMsg["URL"] = null;
                    responseMsg["isSuccess"] = false;
                    responseMsg["msg"] = "키가 유효하지 않습니다.";
                    return;
                }

                ReturnValue status = null;
                bool PAGINGNUM = false;
                bool APPVISIBLE = true;
                string fileName = requestMsg["FileName"].ToString(); // 파일 이름
                string convertIMG = requestMsg["ConvertIMG"].ToString(); // 0: NONE  1:JPEG  2:PNG  3:BMP
                //string docPassword = requestMsg["DocPassword"].ToString(); // 문서 비밀번호
                string docPassword = null; // 문서 비밀번호
                string dataPath = Properties.Settings.Default.DataPath; // 파일 출력경로

                string documents = @"workspace";
                string tmpPath = @"tmp";

                string md5_filechecksum = MD5_CheckSUM(dataPath + @"\" + tmpPath + @"\" + fileName).ToString(); // 파일 체크섬 추출

                DirectoryInfo createDirectory = new DirectoryInfo(dataPath + @"\" + documents + @"\" + md5_filechecksum);
                FileInfo moveFile = new FileInfo(dataPath + @"\" + tmpPath + @"\" + fileName);

                string fileFullPath = createDirectory.FullName + @"\" + fileName;
                string outPath = Path.GetDirectoryName(fileFullPath) + @"\" + Path.GetFileNameWithoutExtension(fileName) + ".pdf";

                // 디렉토리 생성
                if (!createDirectory.Exists)
                    createDirectory.Create();
                // 파일 이동
                //if (moveFile.Exists)
                //   moveFile.MoveTo(fileFullPath);

                // PDF로 변환
                if (Path.GetExtension(fileFullPath).Equals(".docx") || Path.GetExtension(fileFullPath).Equals(".doc"))
                {
                    status = WordConvert_Core.WordSaveAs(fileFullPath, outPath, docPassword, PAGINGNUM, APPVISIBLE);
                }
                else if (Path.GetExtension(fileFullPath).Equals(".xlsx") || Path.GetExtension(fileFullPath).Equals(".xls"))
                {
                    status = ExcelConvert_Core.ExcelSaveAs(fileFullPath, outPath, docPassword, PAGINGNUM, APPVISIBLE);
                }
                else if (Path.GetExtension(fileFullPath).Equals(".pptx") || Path.GetExtension(fileFullPath).Equals(".ppt"))
                {
                    status = PowerPointConvert_Core.PowerPointSaveAs(fileFullPath, outPath, docPassword, PAGINGNUM, APPVISIBLE);
                }
                else if (Path.GetExtension(fileFullPath).Equals(".hwp"))
                {
                    status = HWPConvert_Core.HwpSaveAs(fileFullPath, outPath, PAGINGNUM);
                }
                else if (Path.GetExtension(fileFullPath).Equals(".pdf"))
                {
                    ReturnValue pdfreturnValue = new ReturnValue();
                    pdfreturnValue.isSuccess = true;
                    pdfreturnValue.Message = "PDF파일은 변환할 필요가 없습니다.";
                    pdfreturnValue.PageCount = ConvertImg.pdfPageCount(fileFullPath);
                    status = pdfreturnValue;
                }
                else
                {
                    responseMsg["msg"] = "[상태]   지원포맷 아님. 파싱한 확장자: " + Path.GetExtension(fileFullPath);
                    responseMsg["isSuccess"] = false;
                    return;
                }

                if (!convertIMG.Equals("0"))
                {
                    // 이미지로 변환
                    ReturnValue pdfToImgReturn = null;
                    String imageOutput = Path.GetDirectoryName(outPath) + "\\" + Path.GetFileNameWithoutExtension(outPath) + "\\";
                    if (!new DirectoryInfo(imageOutput).Exists)
                        new DirectoryInfo(imageOutput).Create();
                    if (convertIMG.Equals("1"))
                    {
                        pdfToImgReturn = ConvertImg.PDFtoJpeg(outPath, imageOutput);
                    }
                    else if (convertIMG.Equals("2"))
                    {
                        pdfToImgReturn = ConvertImg.PDFtoPng(outPath, imageOutput);
                    }
                    else if (convertIMG.Equals("3"))
                    {
                        pdfToImgReturn = ConvertImg.PDFtoBmp(outPath, imageOutput);
                    }
                    if (pdfToImgReturn.isSuccess)
                    {
                        responseMsg["convertImgCnt"] = string.Format("{0}", pdfToImgReturn.PageCount);
                    }
                    else
                    {
                        responseMsg["convertImgCnt"] = -1;
                    }
                }


                responseMsg["URL"] = Properties.Settings.Default.serverIP + ":" + Properties.Settings.Default.fileServerPORT + "/" + documents + "/" + md5_filechecksum;
                responseMsg["isSuccess"] = status.isSuccess;
                if(PAGINGNUM)
                    responseMsg["pageNum"] = status.PageCount;
                responseMsg["msg"] = status.Message;

            } catch(Exception e1)
            {
                responseMsg["isSuccess"] = false;
                responseMsg["msg"] = e1.Message;
            }

            DevLog.Write(string.Format("\r\n[Server => Client]\r\n{0}\r\n", responseMsg.ToString()), LOG_LEVEL.INFO); // 클라이언트가 서버로 보낸 메시지

            List<byte> dataSource = new List<byte>();
            dataSource.AddRange(BitConverter.GetBytes((int)PACKETID.REQ_ECHO));
            dataSource.AddRange(BitConverter.GetBytes(requestInfo.Body.Length));
            dataSource.AddRange(Encoding.Unicode.GetBytes(responseMsg.ToString()));

            session.Send(dataSource.ToArray(), 0, dataSource.Count);
        }

        private string MD5_CheckSUM(string filename)
        {
            using (var fs = File.OpenRead(filename))
            using (var md5 = new MD5CryptoServiceProvider())
                return string.Join("", md5.ComputeHash(fs).ToArray().Select(i => i.ToString("X2")));
        }
    }

    public class PK_ECHO
    {
        public string msg;
    }
}
