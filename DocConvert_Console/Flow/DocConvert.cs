using DocConvert_Console.Common;
using DocConvert_Core.FileLib;
using DocConvert_Core.OfficeLib;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace DocConvert_Console.Flow
{
    class DocConvert
    {
        /// <summary>
        /// 문서 변환로직
        /// </summary>
        /// <param name="requestMsg">요청 JSON 데이터</param>
        /// <returns>응답 JSON 데이터</returns>
        public JObject _DocConvert(string fileName, string docPassword, string drmUseYn, string drmType)
        {
            LogMgr.Write("Method DocConvert Start ==>");

            JObject responseMsg = new JObject();
            bool PAGINGNUM = false; // 변환 전 페이징 체크 여부 (느려짐)
            bool APPVISIBLE = false; // 변환 진행상황 표시 여부

            DirectoryInfo workspacePath = new DirectoryInfo(Program.IniProperties.Workspace_Directory); // 파일 출력경로
            DirectoryInfo tmpPath = new DirectoryInfo(workspacePath + @"\tmp");
            FileInfo srcFile = new FileInfo(tmpPath.FullName + @"\" + fileName);
            DirectoryInfo dataTodayMD5Path = new DirectoryInfo(workspacePath + @"\data\" + DateTime.Now.ToString("yyyyMMdd") + @"\" + MD5_CheckSUM(srcFile.FullName));
            FileInfo targetFile = new FileInfo(dataTodayMD5Path.FullName + @"\" + fileName);
            FileInfo newPdfFile = new FileInfo(dataTodayMD5Path.FullName + @"\" + Path.GetFileNameWithoutExtension(fileName));

            LogMgr.Write("Initialized variable ========>", LOG_LEVEL.DEBUG);
            LogMgr.Write("workspacePath ===> " + workspacePath.FullName, LOG_LEVEL.DEBUG);
            LogMgr.Write("tmpPath ===> " + tmpPath.FullName, LOG_LEVEL.DEBUG);
            LogMgr.Write("srcFile ===> " + srcFile.FullName, LOG_LEVEL.DEBUG);
            LogMgr.Write("dataTodayMD5Path ===> " + dataTodayMD5Path.FullName, LOG_LEVEL.DEBUG);
            LogMgr.Write("targetFile ===> " + targetFile.FullName, LOG_LEVEL.DEBUG);
            LogMgr.Write("=============================>", LOG_LEVEL.DEBUG);

            // 경로에 파일이 있는지 체크
            if (!srcFile.Exists)
            {
                throw new IOException("There is no file in " + srcFile.FullName);
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
                srcFile.MoveTo(targetFile.FullName);
                LogMgr.Write(srcFile.ToString() + " Move file to " + targetFile.FullName);
            }

            // PDF로 변환
            if (Path.GetExtension(targetFile.FullName).Equals(".docx") || Path.GetExtension(targetFile.FullName).Equals(".doc") || Path.GetExtension(targetFile.FullName).Equals(".txt") || Path.GetExtension(targetFile.FullName).Equals(".html"))
            {
                WordConvert_Core.WordSaveAs(targetFile.FullName, newPdfFile.FullName, docPassword, PAGINGNUM, APPVISIBLE);
            }
            else if (Path.GetExtension(targetFile.FullName).Equals(".xlsx") || Path.GetExtension(targetFile.FullName).Equals(".xls") || Path.GetExtension(targetFile.FullName).Equals(".csv"))
            {
                ExcelConvert_Core.ExcelSaveAs(targetFile.FullName, newPdfFile.FullName, docPassword, PAGINGNUM, APPVISIBLE);
            }
            else if (Path.GetExtension(targetFile.FullName).Equals(".pptx") || Path.GetExtension(targetFile.FullName).Equals(".ppt"))
            {
                PowerPointConvert_Core.PowerPointSaveAs(targetFile.FullName, newPdfFile.FullName, docPassword, PAGINGNUM, APPVISIBLE);
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
