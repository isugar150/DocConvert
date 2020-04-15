using DocConvert_Core.imageLib;
using DocConvert_Core.interfaces;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DocConvert_Core.WebCaptureLib
{
    public class WebCapture_Core
    {
        /// <summary>
        /// 참고 API https://phantomjs.org/
        /// </summary>
        /// <param name="Url">캡쳐할 웹사이트의 페이지</param>
        /// <param name="outPath">내보낼 경로</param>
        /// <returns></returns>
        [STAThread]
        public static ReturnValue WebCapture(string Url, string outPath)
        {
            ReturnValue returnValue = new ReturnValue();
            try
            {
                string phantomJSPath = Application.StartupPath + @"\phantomjs.exe";
                string optionJS = Application.StartupPath + @"\rasterize.js";

                string arguments = string.Format("{0} {1} {2}", "\"" + optionJS + "\"", "\"" + Url + "\"", "\"" + outPath + @"\" + new Uri(Url).Authority + ".png" + "\"");

                Process process = new Process();
                ProcessStartInfo processStartInfo = new ProcessStartInfo();
                processStartInfo.FileName = phantomJSPath;
                processStartInfo.Arguments = arguments;

                process.StartInfo = processStartInfo;
                process.Start();
                process.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                // 20초동안 끝나지않았을때 강제종료
                DateTime timeTaken = DateTime.Now.AddSeconds(20);
                while (!process.HasExited)
                {
                    if(DateTime.Now > timeTaken)
                        process.Kill();
                    Thread.Sleep(300);
                }
                process.Dispose();

                if(new FileInfo(outPath + @"\" + new Uri(Url).Authority + ".png").Exists)
                {
                    returnValue.isSuccess = true;
                    returnValue.Message = "WebCapture에 성공하였습니다.";
                    returnValue.PageCount = 1;
                }
                else
                {
                    returnValue.isSuccess = false;
                    returnValue.Message = "WebCapture에 실패하였습니다.";
                    returnValue.PageCount = 0;
                }
            }
            catch (Exception e1)
            {
                returnValue.isSuccess = false;
                returnValue.Message = e1.Message;
                returnValue.PageCount = 0;
            }

            return returnValue;
        }
    }
}
