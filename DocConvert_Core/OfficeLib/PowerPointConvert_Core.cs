using DocConvert_Core.FileLib;
using DocConvert_Core.interfaces;
using Microsoft.Office.Interop.PowerPoint;
using NLog;
using System;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using MsoTriState = Microsoft.Office.Core.MsoTriState;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

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
        public static ReturnValue PowerPointSaveAs(string FilePath, string outPath, string docPassword, bool pageCounting, bool appvisible)
        {
            ReturnValue returnValue = new ReturnValue();
            logger.Info("==================== Start ====================");
            logger.Info("Method: " + MethodBase.GetCurrentMethod().Name + ", FilePath: " + FilePath + ", outPath: " + outPath + ", docPassword: " + docPassword);
            #region File Unlock
            try
            {
                LockFile.UnLock_File(FilePath);
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
                // 파워포인트 매크로 실행 비활성화
                powerpoint.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;
                // 파워포인트 알림 비활성화
                powerpoint.DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsNone;
                #endregion

                #region 열기 옵션 https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2010/ff763759(v%3Doffice.14)
                Presentation doc;
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
                MsoTriState OpenAndRepair = MsoTriState.msoTrue; // 200y Only
                #endregion
                #region 문서 열기
                try // Presentations.Open https://docs.microsoft.com/en-us/office/vba/api/powerpoint.presentations.open
                {
                    doc = multiPresentations.Open(
                        FilePath,
                        ReadOnly,
                        Untitled,
                        WithWindow
                    );
                }
                catch (Exception) // Presentations.Open2007 https://docs.microsoft.com/en-us/office/vba/api/powerpoint.presentations.open2007
                {
                    doc = multiPresentations.Open2007(
                        FilePath,
                        ReadOnly,
                        Untitled,
                        WithWindow,
                        OpenAndRepair
                    );
                }
                #endregion
                #region 페이지수 얻기
                if (pageCounting)
                {
                    Slides slides = null;
                    try
                    {
                        slides = doc.Slides;
                        returnValue.PageCount = slides.Count;
                    }
                    catch (Exception e1)
                    {
                        returnValue.PageCount = -1;
                        logger.Error("페이지 카운트 가져오는중 오류발생");
                        logger.Error("오류내용: " + e1.Message);
                    }
                }
                #endregion
                #region 저장 옵션 https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2010/ff762466(v%3Doffice.14)
                PpSaveAsFileType ppSaveAsFileType = PpSaveAsFileType.ppSaveAsPDF;
                MsoTriState msoTriState = MsoTriState.msoFalse;
                #endregion
                #region PDF저장
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
                returnValue.isSuccess = true;
                returnValue.Message = "변환에 성공하였습니다.";
                return returnValue;
            }
            catch (Exception e1)
            {
                logger.Error("======= Method: " + MethodBase.GetCurrentMethod().Name + " =======");
                logger.Error(new StackTrace(e1, true).ToString());
                logger.Error("변환 실패: " + e1.Message);
                logger.Error("================ End ================");
                throw e1;
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
