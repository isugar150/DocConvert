using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

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
        public static bool HwpSaveAs(String FilePath, String outPath)
        {
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

            try
            {
                AxHWPCONTROLLib.AxHwpCtrl axHwpCtrl = new AxHWPCONTROLLib.AxHwpCtrl();
                axHwpCtrl.CreateControl();
                axHwpCtrl.Enabled = true;
                axHwpCtrl.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample");
                #region 문서 열기
                if(axHwpCtrl.Open(FilePath, "HWP", "suspendpassword:TRUE;forceopen:TRUE;versionwarning:FALSE"))
                {
                    #region PDF저장
                    axHwpCtrl.SaveAs(outPath, "PDF", "");
                    #endregion
                    #region 문서 닫기
                    axHwpCtrl.Clear();
                    #endregion
                    logger.Info("변환 성공");
                    return true;
                }
                else
                {
                    logger.Info("한글 파일 ");
                    return false;
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
                return false;
            }

            finally
            {
                #region 앱 종료
                #endregion
                logger.Info("==================== End ====================");
            }
        }
    }
}
