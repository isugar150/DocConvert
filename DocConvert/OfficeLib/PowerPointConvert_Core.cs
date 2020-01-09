﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocConvert_Core.FileLib;
using Microsoft.Office.Interop.PowerPoint;
using NetOffice.OfficeApi.Enums;
using NLog;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using MsoTriState = Microsoft.Office.Core.MsoTriState;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Diagnostics;
using System.Threading;

namespace DocConvert_Core.OfficeLib
{
    public class PowerPointConvert_Core
    {
        private static Logger logger = LogManager.GetLogger("DocConvert_Engine_Log");
        /// <summary>
        /// 워드파일을 PDF로 변환
        /// </summary>
        /// <param name="FilePath">소스파일</param>
        /// <param name="outPath">저장파일</param>
        /// <param name="docPassword">문서 비밀번호</param>
        /// <returns></returns>
        public static bool PowerPointSaveAs(String FilePath, String outPath, String docPassword, bool appvisible)
        {
            logger.Info("==================== Start ====================");
            logger.Info("Method: " + MethodBase.GetCurrentMethod().Name + ", FilePath: " + FilePath + ", outPath: " + outPath + ", docPassword: " + docPassword);
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
            _Application powerpoint = new Application();
            try
            {
                Presentations multiPresentations = powerpoint.Presentations;
                #region 앱 옵션
                powerpoint.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable; // 매크로 실행 안되게
                powerpoint.DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsNone;
                #endregion

                #region 열기 옵션 https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2010/ff763759(v%3Doffice.14)
                MsoTriState ReadOnly = MsoTriState.msoTrue;
                MsoTriState Untitled = MsoTriState.msoFalse;
                MsoTriState WithWindow;
                if (appvisible)
                {
                    WithWindow = MsoTriState.msoTrue;
                }
                else
                {
                    WithWindow = MsoTriState.msoFalse;
                }
                #endregion
                #region 문서 열기
                Presentation doc = multiPresentations.Open(
                    FilePath,
                    ReadOnly,
                    Untitled,
                    WithWindow
                );
                #endregion
                #region 저장 옵션 https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2010/ff762466(v%3Doffice.14)
                #endregion
                #region PDF저장
                PpSaveAsFileType ppSaveAsFileType = PpSaveAsFileType.ppSaveAsPDF;
                MsoTriState msoTriState = MsoTriState.msoFalse;
                doc.SaveAs(
                    outPath,
                    ppSaveAsFileType,
                    msoTriState
                );
                #endregion
                #region 문서 닫기
                doc.Close();
                #endregion
                logger.Info("변환 성공");
                return true;
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
                powerpoint.Quit();
                Marshal.ReleaseComObject(powerpoint);
                powerpoint = null;
                powerpoint = new Application { };
                powerpoint.Quit();
                Marshal.ReleaseComObject(powerpoint);
                powerpoint = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                #endregion
                logger.Info("==================== End ====================");
            }
        }
    }
}
