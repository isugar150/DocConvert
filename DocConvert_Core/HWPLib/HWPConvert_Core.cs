using DocConvert_Core.FileLib;
using DocConvert_Core.interfaces;
using NLog;
using System;
using System.Reflection;
using System.Threading;

namespace DocConvert_Core.HWPLib
{
    public class HWPConvert_Core
    {
        private static Logger logger = LogManager.GetLogger("DocConvert_Engine_Log");
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
                logger.Info("파일 언락 성공!");
            }
            catch (Exception e1)
            {
                logger.Info("파일 언락 실패! 자세한내용 로그 참고");
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
                            logger.Error("페이지 카운트 가져오는중 오류발생");
                            logger.Error("오류내용: " + e1.Message);
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
                    returnValue.Message = "변환에 성공하였습니다.";
                    return returnValue;
                }
                else
                {
                    returnValue.isSuccess = false;
                    returnValue.Message = "문서 변환에 실패하였습니다. (알수없는 오류)";
                    return returnValue;
                }
                #endregion
            }

            catch (Exception e1)
            {
                try
                {
                    throw e1;
                }
                catch (Exception)
                {
                    returnValue.isSuccess = false;
                    returnValue.Message = "문서 변환에 실패하였습니다. (변환중 다른요청이 들어왔습니다.)";
                    return returnValue;
                }
            }

            finally
            {
                #region 앱 종료
                axHwpCtrl.Dispose();
                #endregion
                logger.Info("==================== End ====================");
            }
        }
    }
}
