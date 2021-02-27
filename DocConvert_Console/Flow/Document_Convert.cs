using DocConvert_Console.Common;
using DocConvert_Core.FileLib;
using DocConvert_Core.HWPLib;
using DocConvert_Core.imageLib;
using DocConvert_Core.interfaces;
using DocConvert_Core.OfficeLib;
using DocConvert_Core.ZipLib;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace DocConvert_Console.Flow
{
    public class Document_Convert
    {
        public JObject docConvert(string requestInfo)
        {
            ReturnValue status = new ReturnValue();
            JObject responseMsg = new JObject();
            JObject requestMsg = new JObject();

            string req_Key = "";
            string req_FileData = "";
            string req_Method = "";
            string req_FileName = "";
            int req_ConvertIMG = 0;
            string req_DrmUseYn = "";
            string req_DrmType = "";

            string res_FileData = "";
            bool res_isSuccess = false;
            string res_Message = "";
            string res_Method = "";
            int res_ImgCnt = 0;
            try
            {
                requestMsg = JObject.Parse(requestInfo); // 요청받은 JSON 파싱

                req_Key = requestMsg["KEY"]?.ToString() ?? ""; // 신뢰하는 클라이언트 확인.
                req_FileData = requestMsg["FileData"]?.ToString() ?? ""; // 물리 파일정보
                req_Method = requestMsg["Method"]?.ToString() ?? ""; // 변환 메소드
                req_FileName = requestMsg["FileName"]?.ToString() ?? ""; // 요청한 파일 이름.
                req_ConvertIMG = int.Parse(requestMsg["ConvertIMG"]?.ToString() ?? "0"); // 이미지 변환 타입  0: NONE  1:JPEG  2:PNG  3:BMP
                req_DrmUseYn = requestMsg["DRM_useYn"]?.ToString() ?? ""; // DRM 사용 여부
                req_DrmType = requestMsg["DRM_Type"]?.ToString() ?? ""; // DRM 타입 

                LogMgr.Write("=======[Client => Server]=======");
                LogMgr.Write("req_Key ==> " + req_Key);
                LogMgr.Write("req_Method ==> " + req_Method);
                LogMgr.Write("req_FileName ==> " + req_FileName);
                LogMgr.Write("req_ConvertIMG ==> " + req_ConvertIMG);
                LogMgr.Write("============================");
                res_Method = req_Method;

                if (req_Method.Equals("DocConvert"))
                {
                    #region DocConvert
                    bool PAGINGNUM = false; // 변환 전 페이징 체크 여부 (느려짐)
                    bool APPVISIBLE = true; // 변환 진행상황 표시 여부
                    //string docPassword = requestMsg["DocPassword"].ToString();
                    string docPassword = null; // 문서 비밀번호
                    string dataPath = Program.IniProperties.Workspace_Directory; // 파일 출력경로

                    string documents = @"workspace";
                    string tmpDirName = @"tmp";

                    string tmpPath = dataPath + @"\" + tmpDirName + @"\" + req_FileName;

                    string md5_filechecksum = MD5_CheckSUM(tmpPath).ToString(); // 파일 체크섬 추출

                    DirectoryInfo createDirectory = new DirectoryInfo(dataPath + @"\" + documents + @"\" + md5_filechecksum);
                    FileInfo moveFile = new FileInfo(tmpPath);

                    string fileFullPath = createDirectory.FullName + @"\" + req_FileName;
                    string outPath = Path.GetDirectoryName(fileFullPath) + @"\" + Path.GetFileNameWithoutExtension(req_FileName) + ".pdf";

                    // 폴더가 있으면 삭제
                    if (createDirectory.Exists)
                    {
                        createDirectory.Delete(true);
                        LogMgr.Write("다음 폴더가 이미 존재해서 삭제 후 다시만듬: " + createDirectory.ToString());
                    }

                    createDirectory.Create();

                    #region File Unlock
                    try
                    {
                        LockFile.UnLock_File(moveFile.ToString());
                    }
                    catch (Exception e1)
                    {
                        LogMgr.Write("파일 언락 실패! 자세한내용 로그 참고");
                        LogMgr.Write(e1.Message);
                    }
                    #endregion

                    // 파일 이동
                    if (moveFile.Exists)
                    {
                        moveFile.MoveTo(fileFullPath);
                        LogMgr.Write(tmpPath.ToString() + " 파일을 " + fileFullPath + "로 이동.");
                    }

                    // DRM 변환
                    bool DRM_useYn = Program.IniProperties.DRM_useYn;

                    LogMgr.Write("DRM 사용여부: " + DRM_useYn);
                    // DRM 사용 여부
                    if (DRM_useYn && req_DrmUseYn.Equals("True"))
                    {
                        LogMgr.Write("=================== DRM ===================");
                        string[] DRM_args_ori = Program.IniProperties.DRM_Args.Split(',');
                        string[] DRM_args = new string[DRM_args_ori.Length];

                        FileInfo drmFile = new FileInfo(Path.GetDirectoryName(fileFullPath) + @"\ori_" + Path.GetFileName(req_FileName));
                        FileInfo decFile = new FileInfo(fileFullPath);

                        if (decFile.Exists && req_DrmUseYn.Equals("True"))
                        {
                            LogMgr.Write("DRM 파일 이름 변경 ==> " + drmFile.FullName);
                            File.Move(decFile.FullName, drmFile.FullName);
                        }

                        // DRM 아규먼트 
                        for (int i = 0; i < DRM_args_ori.Length; i++)
                        {
                            if (DRM_args_ori[i].Equals("$Full_Path$"))
                                DRM_args[i] = drmFile.FullName;
                            else if (DRM_args_ori[i].Equals("$File_Path$"))
                                DRM_args[i] = Path.GetDirectoryName(drmFile.FullName) + @"\";
                            else if (DRM_args_ori[i].Equals("$File_Name$"))
                                DRM_args[i] = Path.GetFileName(drmFile.FullName);
                            else if (DRM_args_ori[i].Equals("$Out_Full_Path$"))
                                DRM_args[i] = decFile.FullName;
                            else if (DRM_args_ori[i].Equals("$DRM_Type$"))
                                DRM_args[i] = requestMsg["DRM_Type"].ToString();
                            else
                                DRM_args[i] = DRM_args_ori[i];
                        }

                        string excuteFile = Program.IniProperties.DRM_Path;
                        string DRM_arguments = "";
                        for (int i = 0; i < DRM_args.Length; i++)
                        {
                            if (i == 0)
                                DRM_arguments = "\"" + DRM_args[i] + "\"";
                            else
                                DRM_arguments += " \"" + DRM_args[i] + "\"";
                        }
                        LogMgr.Write(excuteFile + " " + DRM_arguments);

                        // DRM 실행
                        string result = ProcessUtil.RunProcess(excuteFile, DRM_arguments).Replace(Environment.NewLine, "");

                        LogMgr.Write("DRM Result ===> " + result);

                        if (result.Equals(Program.IniProperties.DRM_Result))
                            LogMgr.Write("DRM 변환 성공 !!");
                        else
                        {
                            LogMgr.Write("DRM 변환 실패");
                            res_FileData = null;
                            res_isSuccess = false;
                            res_Message = "DRM 변환 실패 >> DRM Result: " + result;

                            goto error;
                        }
                        LogMgr.Write("===========================================");
                    }

                    // PDF로 변환
                    if (Path.GetExtension(fileFullPath).Equals(".docx") || Path.GetExtension(fileFullPath).Equals(".doc") || Path.GetExtension(fileFullPath).Equals(".txt") || Path.GetExtension(fileFullPath).Equals(".html"))
                    {
                        status = WordConvert_Core.WordSaveAs(fileFullPath, outPath, docPassword, PAGINGNUM, APPVISIBLE);
                    }
                    else if (Path.GetExtension(fileFullPath).Equals(".xlsx") || Path.GetExtension(fileFullPath).Equals(".xls") || Path.GetExtension(fileFullPath).Equals(".csv"))
                    {
                        status = ExcelConvert_Core.ExcelSaveAs(fileFullPath, outPath, docPassword, PAGINGNUM, APPVISIBLE);
                    }
                    else if (Path.GetExtension(fileFullPath).Equals(".pptx") || Path.GetExtension(fileFullPath).Equals(".ppt"))
                    {
                        status = PowerPointConvert_Core.PowerPointSaveAs(fileFullPath, outPath, docPassword, PAGINGNUM, APPVISIBLE);
                    }
                    else if (Path.GetExtension(fileFullPath).Equals(".hwp"))
                    {
                        while (Program.isHwpConverting)
                        {
                            Thread.Sleep(300);
                        }
                        Program.isHwpConverting = true;
                        Thread HWPConvert = new Thread(() =>
                        {
                            status = HWPConvert_Core.HwpSaveAs(fileFullPath, outPath, PAGINGNUM);
                        });
                        HWPConvert.SetApartmentState(ApartmentState.STA);
                        HWPConvert.Start();
                        HWPConvert.Join();

                        Program.isHwpConverting = false;
                    }
                    else if (Path.GetExtension(fileFullPath).Equals(".pdf"))
                    {
                        if (req_ConvertIMG.Equals("0"))
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
                        res_FileData = null;
                        res_isSuccess = false;
                        res_Message = "지원포맷 아님. 파싱한 확장자: " + Path.GetExtension(fileFullPath);
                        res_Method = req_Method;

                        goto error;
                    }

                    // 이미지로 변환
                    String imageOutput = Path.GetDirectoryName(outPath) + "\\" + Path.GetFileNameWithoutExtension(outPath) + "\\";
                    if (!new DirectoryInfo(imageOutput).Exists)
                    {
                        new DirectoryInfo(imageOutput).Create();
                    }

                    if (req_ConvertIMG == 1)
                    {
                        status = ConvertImg.PDFtoJpeg(outPath, imageOutput);
                    }
                    else if (req_ConvertIMG == 2)
                    {
                        status = ConvertImg.PDFtoPng(outPath, imageOutput);
                    }
                    else if (req_ConvertIMG == 3)
                    {
                        status = ConvertImg.PDFtoBmp(outPath, imageOutput);
                    }
                    if (status.isSuccess)
                    {
                        string zipoutPath = Path.GetDirectoryName(outPath) + @"\" + Path.GetFileNameWithoutExtension(outPath) + ".zip";
                        string pdfoutPath = Path.GetDirectoryName(outPath) + @"\" + Path.GetFileNameWithoutExtension(outPath) + ".pdf";
                        if (File.Exists(zipoutPath))
                        {
                            File.Delete(zipoutPath);
                        }

                        if (Directory.Exists(outPath + @"\" + Path.GetFileNameWithoutExtension(outPath)))
                        {
                            Directory.Delete(outPath + @"\" + Path.GetFileNameWithoutExtension(outPath));
                        }

                        if (req_ConvertIMG == 0)
                        {
                            ZipLib.CreateZipFile(new[] { pdfoutPath }, zipoutPath);
                        }
                        else
                        {
                            ZipLib.CreateZipFile(Directory.GetFiles(imageOutput), zipoutPath);
                        }


                        responseMsg["convertImgCnt"] = string.Format("{0}", status.PageCount);

                        //Read file to byte array
                        FileStream stream = File.OpenRead(Path.GetDirectoryName(outPath) + @"\" + Path.GetFileNameWithoutExtension(outPath) + ".zip");
                        byte[] fileBytes = new byte[stream.Length];

                        stream.Read(fileBytes, 0, fileBytes.Length);
                        stream.Close();

                        res_FileData = Convert.ToBase64String(fileBytes);

                        res_isSuccess = status.isSuccess;
                        if (PAGINGNUM)
                        {
                            res_ImgCnt = status.PageCount;
                        }

                        res_Message = status.Message;
                    }
                    else
                    {
                        res_FileData = null;
                        res_isSuccess = status.isSuccess;
                        res_Message = status.Message;
                        res_ImgCnt = status.PageCount;
                    }
                    #endregion
                }
                else
                {
                    res_FileData = null;
                    res_isSuccess = false;
                    res_Message = "메소드가 유효하지 않습니다.";
                }

            error:

                responseMsg["isSuccess"] = res_isSuccess;
                responseMsg["Method"] = res_Method;
                responseMsg["Message"] = res_Message;
                responseMsg["FileData"] = res_FileData;
                responseMsg["ImgCnt"] = res_ImgCnt;
            }
            catch (Exception e1)
            {
                LogMgr.Write("변환중 오류발생 자세한 내용은 오류로그 참고", LOG_LEVEL.ERROR);
                LogMgr.Write("==================== Method: " + MethodBase.GetCurrentMethod().Name + " ====================", LOG_LEVEL.ERROR);
                LogMgr.Write(new StackTrace(e1, true).ToString(), LOG_LEVEL.ERROR);
                LogMgr.Write("변환 실패: " + e1.Message, LOG_LEVEL.ERROR);
                LogMgr.Write("==================== End ====================", LOG_LEVEL.ERROR);
                responseMsg["isSuccess"] = false;
                responseMsg["msg"] = e1.Message;
                responseMsg["Method"] = requestMsg["Method"];
            }

            LogMgr.Write("=======[Server => Client]=======");
            LogMgr.Write("res_isSuccess ==> " + res_isSuccess);
            LogMgr.Write("res_Method ==> " + res_Method);
            LogMgr.Write("res_Message ==> " + res_Message);
            LogMgr.Write("res_ImgCnt ==> " + res_ImgCnt);
            LogMgr.Write("============================");

            return responseMsg;
        }

        private string MD5_CheckSUM(string filename)
        {
            using (FileStream fs = File.OpenRead(filename))
            using (MD5CryptoServiceProvider md5 = new MD5CryptoServiceProvider())
            {
                return string.Join("", md5.ComputeHash(fs).ToArray().Select(i => i.ToString("X2")));
            }
        }
    }
}
