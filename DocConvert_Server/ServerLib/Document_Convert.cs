using DocConvert_Core.FileLib;
using DocConvert_Core.HWPLib;
using DocConvert_Core.imageLib;
using DocConvert_Core.interfaces;
using DocConvert_Core.OfficeLib;
using DocConvert_Core.WebCaptureLib;
using DocConvert_Core.ZipLib;
using Newtonsoft.Json.Linq;
using NLog;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Threading;
using System.Windows.Forms;

namespace DocConvert_Server
{
    public class Document_Convert
    {
        private static Logger logger = LogManager.GetLogger("DocConvert_Server_Log");

        /// <summary>
        /// 소켓, 웹소켓에서 들어온 요청을 처리하는 부분
        /// </summary>
        /// <param name="requestInfo">클라이언트에서 던진 String</param>
        /// <returns></returns>
        public JObject document_Convert(string requestInfo)
        {
            ReturnValue status = new ReturnValue();
            JObject responseMsg = new JObject();
            JObject requestMsg = new JObject();
            try
            {
                requestMsg = JObject.Parse(requestInfo); // 요청받은 JSON 파싱

                if (!requestMsg["KEY"].ToString().Equals(Form1.IniProperties.ClientKEY) || requestMsg["KEY"] == null)
                {
                    responseMsg["URL"] = null;
                    responseMsg["isSuccess"] = false;
                    responseMsg["msg"] = "키가 유효하지 않습니다. 확인 후 다시시도하세요.";
                    responseMsg["Method"] = requestMsg["Method"];
                    return responseMsg;
                }

                if (requestMsg["Method"].ToString().Equals("DocConvert"))
                {
                    if (requestMsg["useCompression"].ToString().Equals("True") && requestMsg["ConvertIMG"].ToString().Equals("0"))
                    {
                        responseMsg["URL"] = null;
                        responseMsg["isSuccess"] = false;
                        responseMsg["msg"] = "이미지 변환을 하지 않으면 압축할 수 없습니다.";
                        responseMsg["Method"] = requestMsg["Method"];
                        return responseMsg;
                    }

                    #region DocConvert
                    bool PAGINGNUM = false;
                    bool APPVISIBLE = Form1.IniProperties.OfficeDebugModeYn;
                    string fileName = requestMsg["FileName"].ToString(); // 파일 이름
                    string convertIMG = requestMsg["ConvertIMG"].ToString(); // 0: NONE  1:JPEG  2:PNG  3:BMP
                    //string docPassword = requestMsg["DocPassword"].ToString();
                    string docPassword = null; // 문서 비밀번호
                    string dataPath = Form1.IniProperties.DataPath; // 파일 출력경로

                    string documents = @"workspace";
                    string tmpPath = @"tmp";

                    string md5_filechecksum = MD5_CheckSUM(dataPath + @"\" + tmpPath + @"\" + fileName).ToString(); // 파일 체크섬 추출

                    DirectoryInfo createDirectory = new DirectoryInfo(dataPath + @"\" + documents + @"\" + md5_filechecksum);
                    FileInfo moveFile = new FileInfo(dataPath + @"\" + tmpPath + @"\" + fileName);

                    string fileFullPath = createDirectory.FullName + @"\" + fileName;
                    string outPath = Path.GetDirectoryName(fileFullPath) + @"\" + Path.GetFileNameWithoutExtension(fileName) + ".pdf";

                    // 폴더가 있으면 삭제
                    if (createDirectory.Exists)
                        createDirectory.Delete(true);

                    createDirectory.Create();

                    // 파일 이동
                    if (moveFile.Exists)
                        moveFile.MoveTo(fileFullPath);

                    // PDF로 변환

                    if (Path.GetExtension(fileFullPath).Equals(".docx") || Path.GetExtension(fileFullPath).Equals(".doc") || Path.GetExtension(fileFullPath).Equals(".txt") || Path.GetExtension(fileFullPath).Equals(".html"))
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
                        while (Form1.isHwpConverting)
                        {
                            Thread.Sleep(300);
                        }
                        Form1.isHwpConverting = true;
                        Thread HWPConvert = new Thread(() =>
                        {
                            status = HWPConvert_Core.HwpSaveAs(fileFullPath, outPath, PAGINGNUM);
                        });
                        HWPConvert.SetApartmentState(ApartmentState.STA);
                        HWPConvert.Start();
                        HWPConvert.Join();

                        Form1.isHwpConverting = false;
                    }
                    else if (Path.GetExtension(fileFullPath).Equals(".pdf"))
                    {
                        if (convertIMG.Equals("0"))
                        {
                            ReturnValue pdfreturnValue = new ReturnValue();
                            pdfreturnValue.isSuccess = true;
                            pdfreturnValue.Message = "PDF파일은 변환할 필요가 없습니다.";
                            pdfreturnValue.PageCount = ConvertImg.pdfPageCount(fileFullPath);
                            status = pdfreturnValue;
                        }
                    }
                    else
                    {
                        responseMsg["URL"] = null;
                        responseMsg["isSuccess"] = false;
                        responseMsg["msg"] = "지원포맷 아님. 파싱한 확장자: " + Path.GetExtension(fileFullPath);
                        responseMsg["Method"] = requestMsg["Method"];
                        responseMsg["useCompression"] = requestMsg["useCompression"];
                        return responseMsg;
                    }

                    if (!convertIMG.Equals("0"))
                    {
                        // 이미지로 변환
                        String imageOutput = Path.GetDirectoryName(outPath) + "\\" + Path.GetFileNameWithoutExtension(outPath) + "\\";
                        if (!new DirectoryInfo(imageOutput).Exists)
                            new DirectoryInfo(imageOutput).Create();
                        if (convertIMG.Equals("1"))
                        {
                            status = ConvertImg.PDFtoJpeg(outPath, imageOutput);
                        }
                        else if (convertIMG.Equals("2"))
                        {
                            status = ConvertImg.PDFtoPng(outPath, imageOutput);
                        }
                        else if (convertIMG.Equals("3"))
                        {
                            status = ConvertImg.PDFtoBmp(outPath, imageOutput);
                        }
                        if (status.isSuccess)
                        {
                            responseMsg["convertImgCnt"] = string.Format("{0}", status.PageCount);

                            if (requestMsg["useCompression"].ToString().Equals("True"))
                            {
                                string zipoutPath = Path.GetDirectoryName(outPath) + @"\" + Path.GetFileNameWithoutExtension(outPath) + ".zip";
                                if (File.Exists(zipoutPath))
                                    File.Delete(zipoutPath);
                                if (Directory.Exists(outPath + @"\" + Path.GetFileNameWithoutExtension(outPath)))
                                    Directory.Delete(outPath + @"\" + Path.GetFileNameWithoutExtension(outPath));
                                ZipLib.CreateZipFile(Directory.GetFiles(imageOutput), zipoutPath);
                            }
                        }
                        else
                        {
                            responseMsg["URL"] = null;
                            responseMsg["isSuccess"] = status.isSuccess;
                            responseMsg["msg"] = status.Message;
                            responseMsg["convertImgCnt"] = status.PageCount;
                        }
                    }

                    if (status.isSuccess)
                    {
                        responseMsg["URL"] = "/" + documents + "/" + md5_filechecksum;
                        responseMsg["isSuccess"] = status.isSuccess;
                        if (PAGINGNUM)
                            responseMsg["pageNum"] = status.PageCount;
                        responseMsg["msg"] = status.Message;
                    }
                    responseMsg["useCompression"] = requestMsg["useCompression"];
                    #endregion
                }
                else if (requestMsg["Method"].ToString().Equals("WebCapture"))
                {
                    string guidPath = Guid.NewGuid().ToString() + "_" + DateTime.Now.ToString("yyyy-MM-dd");
                    string dataPath = @"\workspace\" + guidPath; // 파일 출력경로
                    new DirectoryInfo(Form1.IniProperties.DataPath + dataPath).Create();
                    if(!new DirectoryInfo(@"..\..\..\..\CefSharp.Example\Resources").Exists)
                        new DirectoryInfo(@"..\..\..\..\CefSharp.Example\Resources").Create();
                    if (Form1.IniProperties.ChromiumCaptureYn) //Chromium 캡쳐시
                    {
                        Thread WebCapture = new Thread(() =>
                        {
                            WebBrowser webBrowser = new WebBrowser();
                            webBrowser.CreateControl();
                            webBrowser.ScrollBarsEnabled = false;
                            webBrowser.ScriptErrorsSuppressed = true;

                            webBrowser.Navigate(requestMsg["URL"].ToString());
                            while (webBrowser.ReadyState != WebBrowserReadyState.Complete)
                            {
                                Application.DoEvents();
                            }

                            status = WebCapture_Core.ChromiumCapture(requestMsg["URL"].ToString(), Form1.IniProperties.DataPath + dataPath + @"\0.png", webBrowser.Document.Body.ScrollRectangle.Width + 30, Form1.IniProperties.WebCaptureTimeout);
                            webBrowser.Dispose();
                        });
                        WebCapture.SetApartmentState(ApartmentState.STA);
                        WebCapture.Start();
                        WebCapture.Join();
                        if (status.isSuccess)
                        {
                            responseMsg["URL"] = "/" + "workspace" + "/" + guidPath + "/" + "0.png";
                            responseMsg["isSuccess"] = status.isSuccess;
                            responseMsg["msg"] = status.Message;
                            responseMsg["convertImgCnt"] = status.PageCount;
                        }
                        else
                        {
                            responseMsg["isSuccess"] = status.isSuccess;
                            responseMsg["msg"] = status.Message;
                            responseMsg["convertImgCnt"] = status.PageCount;
                        }
                    }
                    else //PhantomJS 사용시
                    {
                        #region WebCapture
                        Thread WebCapture = new Thread(() =>
                        {
                            status = WebCapture_Core.WebCapture(requestMsg["URL"].ToString(), Form1.IniProperties.DataPath + dataPath, Form1.IniProperties.WebCaptureTimeout);
                        });
                        WebCapture.SetApartmentState(ApartmentState.STA);
                        WebCapture.Start();
                        WebCapture.Join();
                        if (status.isSuccess)
                        {
                            responseMsg["URL"] = "/" + "workspace" + "/" + guidPath + "/" + "0.png";
                            responseMsg["isSuccess"] = status.isSuccess;
                            responseMsg["msg"] = status.Message;
                            responseMsg["convertImgCnt"] = status.PageCount;
                        }
                        else
                        {
                            responseMsg["isSuccess"] = status.isSuccess;
                            responseMsg["msg"] = status.Message;
                            responseMsg["convertImgCnt"] = status.PageCount;
                        }
                    }
                    #endregion
                }
                else
                {
                    responseMsg["URL"] = null;
                    responseMsg["isSuccess"] = false;
                    responseMsg["msg"] = "메소드가 유효하지 않습니다.";
                }

                responseMsg["Method"] = requestMsg["Method"];
            }
            catch (Exception e1)
            {
                logger.Info("변환중 오류발생 자세한 내용은 오류로그 참고");
                logger.Error("==================== Method: " + MethodBase.GetCurrentMethod().Name + " ====================");
                logger.Error(new StackTrace(e1, true));
                logger.Error("변환 실패: " + e1.Message);
                logger.Error("==================== End ====================");
                responseMsg["isSuccess"] = false;
                responseMsg["msg"] = e1.Message;
                responseMsg["Method"] = requestMsg["Method"];
            }

            return responseMsg;
        }

        private string MD5_CheckSUM(string filename)
        {
            using (var fs = File.OpenRead(filename))
            using (var md5 = new MD5CryptoServiceProvider())
                return string.Join("", md5.ComputeHash(fs).ToArray().Select(i => i.ToString("X2")));
        }
    }
}
