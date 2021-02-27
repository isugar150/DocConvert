using DocConvert_Core.FileLib;
using DocConvert_Core.interfaces;
using NLog;
using System;
using System.Diagnostics;
using System.Reflection;
using System.Threading;

namespace DocConvert_Core.HWPLib
{
    public class HWPConvert_Core
    {
        private static Logger logger = LogManager.GetLogger("DocConvert_Core_Log");
        /// <summary>
        /// 한글 문서를 PDF로 변환합니다.
        /// </summary>
        /// <param name="FilePath">소스 경로</param>
        /// <param name="outPath">내보낼 경로</param>
        /// <returns></returns>
        [STAThread]
        public static ReturnValue HwpSaveAs(string FilePath, string outPath, bool PageCounting)
        {
            ReturnValue returnValue = new ReturnValue();

            logger.Info("==================== Start ====================");
            logger.Info("Method: " + MethodBase.GetCurrentMethod().Name + ", FilePath: " + FilePath + ", outPath: " + outPath);
            #region File Unlock
            try
            {
                LockFile.UnLock_File(FilePath);
            }
            catch (Exception e1)
            {
                logger.Info("File unlock failed! See log for details");
                logger.Error(e1.Message);
            }
            #endregion
            AxHWPCONTROLLib.AxHwpCtrl axHwpCtrl = null;
            try
            {
                axHwpCtrl = new AxHWPCONTROLLib.AxHwpCtrl();

                while (axHwpCtrl.InvokeRequired)
                {
                    Thread.Sleep(300);
                }

                axHwpCtrl.CreateControl();

                axHwpCtrl.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample");
                #region 문서 열기
                if (axHwpCtrl.Open(FilePath, "HWP", "suspendpassword:TRUE;forceopen:TRUE;versionwarning:FALSE"))
                {
                    #region 페이지수 얻기
                    if (PageCounting)
                    {
                        try
                        {
                            returnValue.PageCount = axHwpCtrl.PageCount;
                        }
                        catch (Exception e1)
                        {
                            returnValue.PageCount = -1;
                            logger.Error("Error fetching page count");
                            logger.Error("Error detail: " + e1.Message);
                        }
                    }
                    #endregion
                    #region PDF저장
                    axHwpCtrl.SaveAs(outPath, "PDF", "");
                    #endregion
                    #region 문서 닫기
                    axHwpCtrl.Clear(1);
                    #endregion
                    returnValue.isSuccess = true;
                    returnValue.Message = "Conversion was successful.";
                    return returnValue;
                }
                else
                {
                    returnValue.isSuccess = false;
                    returnValue.Message = "Conversion failure (Unknown error)";
                    return returnValue;
                }
                #endregion
            }

            catch (Exception e1)
            {
                logger.Error("======= Method: " + MethodBase.GetCurrentMethod().Name + " =======");
                logger.Error(new StackTrace(e1, true).ToString());
                logger.Error("Image conversion failed. " + e1.Message);
                logger.Error("================ End ================");
                try
                {
                    throw e1;
                }
                catch (Exception)
                {
                    returnValue.isSuccess = false;
                    returnValue.Message = "Document conversion failed. (Another request came in during conversion.)";
                    return returnValue;
                }
            }

            finally
            {
                #region 앱 종료
                axHwpCtrl.Dispose();
                axHwpCtrl = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                #endregion
                logger.Info("==================== End ====================");
            }
        }
    }
}
