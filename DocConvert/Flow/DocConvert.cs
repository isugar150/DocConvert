using DocConvert.Common;
using DocConvert_Core.FileLib;
using DocConvert_Core.HWPLib;
using DocConvert_Core.imageLib;
using DocConvert_Core.interfaces;
using DocConvert_Core.OfficeLib;
using DocConvert_Core.ZipLib;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace DocConvert.Flow
{
    class DocConvert
    {
        /// <summary>
        /// 문서 변환로직
        /// </summary>
        /// <param name="requestMsg">요청 JSON 데이터</param>
        /// <returns>응답 JSON 데이터</returns>
        public JObject _DocConvert(string fileName, string docPassword, int convertImg, string drmUseYn, string drmType)
        {
            LogMgr.Write("Method DocConvert Start ==>");

            JObject responseMsg = new JObject();
            ReturnValue status = new ReturnValue();
            bool PAGINGNUM = false; // 변환 전 페이징 체크 여부 (느려짐)
            bool APPVISIBLE = false; // 변환 진행상황 표시 여부

            DirectoryInfo workspacePath = new DirectoryInfo(Program.IniProperties.Workspace_Directory); // 파일 출력경로
            DirectoryInfo tmpPath = new DirectoryInfo(workspacePath + @"\tmp");
            FileInfo srcFile = new FileInfo(tmpPath.FullName + @"\" + fileName);
            DirectoryInfo dataTodayMD5Path = null;
            try
            {
                dataTodayMD5Path = new DirectoryInfo(workspacePath.FullName + @"\data\" + DateTime.Now.ToString("yyyyMMdd") + @"\" + MD5_CheckSUM(srcFile.FullName));
            } catch(Exception)
            {
                responseMsg["ResultCode"] = define.INVALID_FILE_NOT_FOUND_ERROR.ToString();
                responseMsg["Message"] = "The file does not exist at " + srcFile.FullName;
                return responseMsg;
            }
            FileInfo targetFile = new FileInfo(dataTodayMD5Path.FullName + @"\" + fileName);
            FileInfo newPdfFile = new FileInfo(dataTodayMD5Path.FullName + @"\" + Path.GetFileNameWithoutExtension(fileName) + ".pdf");
            FileInfo newZipFile = new FileInfo(dataTodayMD5Path.FullName + @"\" + Path.GetFileNameWithoutExtension(fileName) + ".zip");
            LogMgr.Write("Initialized variable ========>", LOG_LEVEL.DEBUG);
            LogMgr.Write("workspacePath ===> " + workspacePath.FullName, LOG_LEVEL.DEBUG);
            LogMgr.Write("tmpPath ===> " + tmpPath.FullName, LOG_LEVEL.DEBUG);
            LogMgr.Write("srcFile ===> " + srcFile.FullName, LOG_LEVEL.DEBUG);
            LogMgr.Write("dataTodayMD5Path ===> " + dataTodayMD5Path.FullName, LOG_LEVEL.DEBUG);
            LogMgr.Write("targetFile ===> " + targetFile.FullName, LOG_LEVEL.DEBUG);
            LogMgr.Write("=============================>", LOG_LEVEL.DEBUG);

            // convertImg 변수가 유효하지 않으면 리턴.
            if (convertImg != 0 && convertImg != 1 && convertImg != 2 && convertImg != 3)
            {
                responseMsg["ResultCode"] = define.INVALID_IMAGE_CONVERT_REQUEST_ERROR.ToString();
                responseMsg["Message"] = "Invalid Image Convert Request Error";
                return responseMsg;
            }

            // 이미 생성된 파일이 있으면..
            if (targetFile.Exists)
            {
                // PDF변환이 완료된 상태면..
                if (newPdfFile.Exists)
                {
                    // PDF로 변환을 요청하였을 경우..
                    if (convertImg == 0)
                    {
                        status.isSuccess = true;
                        srcFile.Delete();
                        goto EndPoint;
                    }
                    // 이미지 변환을 요청하였을 경우..
                    else
                    {
                        // 경로에 Zip파일이 있을 경우
                        if (newZipFile.Exists)
                        {
                            status.isSuccess = true;
                            srcFile.Delete();
                            goto EndPoint;
                        }
                        else
                        {
                            srcFile.Delete();
                            goto ConvertImagePoint;
                        }
                    }
                }
            }

            // 오늘 날짜로 된 또는 MD5 폴더가 없으면 생성한다.
            if (!dataTodayMD5Path.Exists)
            {
                dataTodayMD5Path.Create();
                LogMgr.Write("Created DataTodayMD5Path Directory " + dataTodayMD5Path.FullName, ConsoleColor.Yellow, LOG_LEVEL.INFO);
            }

            // 대상이 잠겨있으면 해제한다.
            try
            {
                LockFile.UnLock_File(srcFile.FullName);
            }
            catch (Exception e1)
            {
                LogMgr.Write("File unlock failed! See log for details");
                LogMgr.Write(e1.Message);
            }

            // tmp폴더에 있는 파일을 타겟 폴더로 이동
            if (srcFile.Exists)
            {
                if (targetFile.Exists)
                    targetFile.Delete();
                new FileInfo(srcFile.FullName).MoveTo(targetFile.FullName); // 새로 안만들면 srcFile의 경로가 바뀌어버림
                LogMgr.Write(srcFile.ToString() + " Move file to " + targetFile.FullName);
            }

            // DRM 변환
            bool DRM_useYn = Program.IniProperties.DRM_useYn;
            
            LogMgr.Write("use DRM: " + DRM_useYn);
            // DRM 사용 여부
            if (DRM_useYn && Regex.IsMatch(drmUseYn, "Y", RegexOptions.IgnoreCase))
            {
                LogMgr.Write("=================== DRM ===================");
                string[] DRM_args_ori = Program.IniProperties.DRM_Args.Split(',');
                string[] DRM_args = new string[DRM_args_ori.Length];

                FileInfo drmFile = new FileInfo(dataTodayMD5Path.FullName + @"\ori_" + Path.GetFileName(fileName));
                FileInfo decFile = new FileInfo(targetFile.FullName);

                if (decFile.Exists && Regex.IsMatch(drmUseYn, "Y", RegexOptions.IgnoreCase))
                {
                    LogMgr.Write("DRM file name change ==> " + drmFile.FullName);
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
                        DRM_args[i] = drmType;
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
                    LogMgr.Write("DRM decryption success!");
                else
                {
                    LogMgr.Write("DRM decryption failed!");
                    responseMsg["ResultCode"] = define.DRM_DECRYPTION_ERROR.ToString();
                    responseMsg["Message"] = "DRM decryption failed.";
                    return responseMsg;
                }
                LogMgr.Write("===========================================");
            }

            // PDF로 변환
            if (Path.GetExtension(targetFile.FullName).Equals(".docx") || Path.GetExtension(targetFile.FullName).Equals(".doc") || Path.GetExtension(targetFile.FullName).Equals(".txt") || Path.GetExtension(targetFile.FullName).Equals(".html"))
            {
                status = WordConvert_Core.WordSaveAs(targetFile.FullName, newPdfFile.FullName, docPassword, PAGINGNUM, APPVISIBLE);
            }
            else if (Path.GetExtension(targetFile.FullName).Equals(".xlsx") || Path.GetExtension(targetFile.FullName).Equals(".xls") || Path.GetExtension(targetFile.FullName).Equals(".csv"))
            {
                status = ExcelConvert_Core.ExcelSaveAs(targetFile.FullName, newPdfFile.FullName, docPassword, PAGINGNUM, APPVISIBLE);
            }
            else if (Path.GetExtension(targetFile.FullName).Equals(".pptx") || Path.GetExtension(targetFile.FullName).Equals(".ppt"))
            {
                status = PowerPointConvert_Core.PowerPointSaveAs(targetFile.FullName, newPdfFile.FullName, docPassword, PAGINGNUM, APPVISIBLE);
            }
            else if (Path.GetExtension(targetFile.FullName).Equals(".hwp"))
            {
                while (Program.isHwpConverting)
                {
                    Thread.Sleep(300);
                }
                Program.isHwpConverting = true;
                Thread HWPConvert = new Thread(() =>
                {
                    status = HWPConvert_Core.HwpSaveAs(targetFile.FullName, newPdfFile.FullName, PAGINGNUM);
                });
                HWPConvert.SetApartmentState(ApartmentState.STA);
                HWPConvert.Start();
                HWPConvert.Join();

                Program.isHwpConverting = false;
            }
            else if (Path.GetExtension(targetFile.FullName).Equals(".pdf"))
            {
                if (convertImg == 0)
                {
                    responseMsg["ResultCode"] = define.PDF_TO_PDF_ERROR.ToString();
                    responseMsg["Message"] = "pdf files cannot be converted to pdf.";
                    return responseMsg;
                }
            }

            ConvertImagePoint:
            // 이미지 변환
            if (convertImg != 0)
            {
                DirectoryInfo imageOutPath = new DirectoryInfo(dataTodayMD5Path + @"\" + Path.GetFileNameWithoutExtension(fileName) + "\\");
                if (!imageOutPath.Exists)
                    imageOutPath.Create();
                if (convertImg == 1)
                {
                    status = ConvertImg.PDFtoJpeg(newPdfFile.FullName, imageOutPath.FullName, PdfiumViewer.PdfRenderFlags.ForPrinting);
                }
                else if (convertImg == 2)
                {
                    status = ConvertImg.PDFtoPng(newPdfFile.FullName, imageOutPath.FullName, PdfiumViewer.PdfRenderFlags.ForPrinting);
                }
                else if (convertImg == 3)
                {
                    status = ConvertImg.PDFtoBmp(newPdfFile.FullName, imageOutPath.FullName, PdfiumViewer.PdfRenderFlags.ForPrinting);
                }

                // 압축
                if(status.isSuccess)
                    ZipLib.CreateZipFile(Directory.GetFiles(imageOutPath.FullName), newZipFile.FullName);
            }

            EndPoint: 
            if (status.isSuccess)
            {
                if (convertImg == 0)
                    responseMsg["FilePath"] = newPdfFile.FullName.Replace(workspacePath.FullName, "").Replace(@"\", "/");
                else
                    responseMsg["FilePath"] = newZipFile.FullName.Replace(workspacePath.FullName, "").Replace(@"\", "/");
                responseMsg["PageCnt"] = ConvertImg.pdfPageCount(newPdfFile.FullName);
                responseMsg["ResultCode"] = define.OK.ToString();
                responseMsg["Message"] = "success";
            }
            else
            {
                responseMsg["ResultCode"] = define.UNDEFINE_ERROR;
                responseMsg["Message"] = status.Message;
            }

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
