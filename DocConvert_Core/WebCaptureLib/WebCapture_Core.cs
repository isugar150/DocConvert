using DocConvert_Core.interfaces;
using NLog;
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;

namespace DocConvert_Core.WebCaptureLib
{
    public class WebCapture_Core
    {
        private static Logger logger = LogManager.GetLogger("DocConvert_Engine_Log");
        /// <summary>
        /// 참고 API https://phantomjs.org/
        /// </summary>
        /// <param name="Url">캡쳐할 웹사이트의 페이지</param>
        /// <param name="outPath">내보낼 경로</param>
        /// <returns></returns>
        [STAThread]
        public static ReturnValue WebCapture(string Url, string outPath, int timeOut)
        {
            ReturnValue returnValue = new ReturnValue();
            try
            {
                string phantomJSPath = Application.StartupPath + @"\phantomjs.exe";
                string optionJS = Application.StartupPath + @"\rasterize.js";

                string arguments = string.Format("{0} {1} {2}", "\"" + optionJS + "\"", "\"" + Url + "\"", "\"" + outPath + @"\" + "0.png" + "\"");

                Process process = new Process();
                ProcessStartInfo processStartInfo = new ProcessStartInfo();
                processStartInfo.FileName = phantomJSPath;
                processStartInfo.Arguments = arguments;

                process.StartInfo = processStartInfo;
                process.Start();
                process.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;

                DateTime timeTaken = DateTime.Now.AddSeconds(timeOut);
                while (!process.HasExited)
                {
                    if (DateTime.Now > timeTaken)
                    {
                        process.Kill();
                    }

                    Thread.Sleep(300);
                }
                process.Dispose();

                if (new FileInfo(outPath + @"\" + "0.png").Exists)
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
                logger.Info("변환중 오류발생 자세한 내용은 오류로그 참고");
                logger.Error("==================== Method: " + MethodBase.GetCurrentMethod().Name + "====================");
                logger.Error(new StackTrace(e1, true));
                logger.Error("변환 실패: " + e1.Message);
                logger.Error("==================== End ====================");
            }

            return returnValue;
        }


        [STAThread]
        public static ReturnValue ChromiumCapture(string Url, string outPath, int docWidth, int timeOut)
        {
            ReturnValue returnValue = new ReturnValue();
            try
            {
                if (docWidth < 1366)
                {
                    docWidth = 1366;
                }

                string ChromiumPath = "\"" + Application.StartupPath + @"\CefSharp\CefSharp.OffScreen.Example.exe" + "\"";

                string arguments = string.Format("{0} {1} {2}", "\"" + Url + "\"", "\"" + outPath + "\"", "\"" + docWidth + "\"");

                Debug.WriteLine(ChromiumPath + " " + arguments);

                Process process = new Process();
                ProcessStartInfo processStartInfo = new ProcessStartInfo();
                processStartInfo.FileName = ChromiumPath;
                processStartInfo.Arguments = arguments;

                process.StartInfo = processStartInfo;
                process.Start();
                process.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;

                DateTime timeTaken = DateTime.Now.AddSeconds(timeOut);
                while (!new FileInfo(outPath).Exists)
                {
                    if (DateTime.Now > timeTaken)
                    {
                        process.Kill();
                    }

                    Thread.Sleep(300);
                }
                Thread.Sleep(1000);
                try
                {
                    process.Kill();
                }
                catch (Exception){ }
                process.Dispose();

                if (new FileInfo(outPath).Exists)
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
                logger.Info("변환중 오류발생 자세한 내용은 오류로그 참고");
                logger.Error("==================== Method: " + MethodBase.GetCurrentMethod().Name + "====================");
                logger.Error(new StackTrace(e1, true));
                logger.Error("변환 실패: " + e1.Message);
                logger.Error("==================== End ====================");
            }

            return returnValue;
        }
    }
}
