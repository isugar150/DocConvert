using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocConvert_Core.interfaces;
using DocConvert_Core.FileLib;
using Microsoft.Win32;
using NLog;

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
        public static ReturnValue HwpSaveAs(String FilePath, String outPath)
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
                
                axHwpCtrl.CreateControl();
                
                axHwpCtrl.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample");
                #region 문서 열기
                if(axHwpCtrl.Open(FilePath, "HWP", "suspendpassword:TRUE;forceopen:TRUE;versionwarning:FALSE"))
                {
                    #region 페이지수 얻기
                    try
                    {
                        returnValue.PageCount = axHwpCtrl.PageCount;
                    } catch (Exception e1)
                    {
                        returnValue.PageCount = -1;
                        logger.Error("페이지 카운트 가져오는중 오류발생");
                        logger.Error("오류내용: " + e1.Message);
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
                logger.Info("변환중 오류발생 자세한 내용은 오류로그 참고");
                logger.Error("==================== Method: " + MethodBase.GetCurrentMethod().Name + " ====================");
                logger.Error(new StackTrace(e1, true));
                logger.Error("변환 실패: " + e1.Message);
                logger.Error("==================== End ====================");
                returnValue.isSuccess = false;
                returnValue.Message = e1.Message;
                return returnValue;
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
