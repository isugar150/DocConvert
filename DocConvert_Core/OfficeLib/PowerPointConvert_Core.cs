﻿using DocConvert_Core.FileLib;
using DocConvert_Core.interfaces;
using Microsoft.Office.Interop.PowerPoint;
using NLog;
using System;
using System.Diagnostics;
using System.Drawing.Printing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using MsoTriState = Microsoft.Office.Core.MsoTriState;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace DocConvert_Core.OfficeLib
{
    public class PowerPointConvert_Core
    {
        private static Logger logger = LogManager.GetLogger("DocConvert_Core_Log");
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
                logger.Info("File unlock failed! See log for details");
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
                powerpoint.DisplayAlerts = PpAlertLevel.ppAlertsNone;
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
                MsoTriState OpenAndRepair = MsoTriState.msoTrue; // 07 Only
                #endregion
                #region 문서 열기
                if (Path.GetExtension(FilePath).Equals(".pptx")) // Presentations.Open https://docs.microsoft.com/en-us/office/vba/api/powerpoint.presentations.open
                {
                    doc = multiPresentations.Open(
                        FileName: FilePath,
                        ReadOnly: ReadOnly,
                        Untitled: Untitled,
                        WithWindow: WithWindow
                    );
                }
                else // Presentations.Open2007 https://docs.microsoft.com/en-us/office/vba/api/powerpoint.presentations.open2007
                {
                    doc = multiPresentations.Open2007(
                        FileName: FilePath,
                        ReadOnly: ReadOnly,
                        Untitled: Untitled,
                        WithWindow: WithWindow,
                        OpenAndRepair: OpenAndRepair
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
                #region 저장 옵션
                // 다른 이름으로 저장 https://docs.microsoft.com/en-us/office/vba/api/powerpoint.presentation.saveas
                PpSaveAsFileType ppSaveAsFileType = PpSaveAsFileType.ppSaveAsPDF;
                MsoTriState msoTriState = MsoTriState.msoFalse;

                // 내보내기 https://docs.microsoft.com/en-us/office/vba/api/powerpoint.presentation.exportasfixedformat
                PpFixedFormatType FixedFormatType = PpFixedFormatType.ppFixedFormatTypePDF;
                PpFixedFormatIntent Intent = PpFixedFormatIntent.ppFixedFormatIntentPrint;
                MsoTriState FrameSlides = MsoTriState.msoFalse;
                PpPrintHandoutOrder HandoutOrder = PpPrintHandoutOrder.ppPrintHandoutVerticalFirst;
                PpPrintOutputType OutputType = PpPrintOutputType.ppPrintOutputSlides;
                MsoTriState PrintHiddenSlides = MsoTriState.msoFalse;
                PowerPoint.PrintRange PrintRange = null;
                PpPrintRangeType RangeType = PpPrintRangeType.ppPrintAll;
                string SlideShowName = string.Empty;
                bool IncludeDocProperties = false;
                bool KeepIRMSettings = false;
                bool DocStructureTags = true;
                bool BitmapMissingFonts = true;
                bool UseISO19005_1 = false;
                #endregion
                #region PDF저장
                try 
                {
                    doc.SaveAs(
                        FileName: outPath,
                        FileFormat: ppSaveAsFileType,
                        EmbedTrueTypeFonts: msoTriState
                    );
                }
                catch (Exception)
                {
                    doc.ExportAsFixedFormat(
                        Path: outPath,
                        FixedFormatType: FixedFormatType,
                        Intent: Intent,
                        FrameSlides: FrameSlides,
                        HandoutOrder: HandoutOrder,
                        OutputType: OutputType,
                        PrintHiddenSlides: PrintHiddenSlides,
                        PrintRange: PrintRange,
                        RangeType: RangeType,
                        SlideShowName: SlideShowName,
                        IncludeDocProperties: IncludeDocProperties,
                        KeepIRMSettings: KeepIRMSettings,
                        DocStructureTags: DocStructureTags,
                        BitmapMissingFonts: BitmapMissingFonts,
                        UseISO19005_1: UseISO19005_1,
                        ExternalExporter: Type.Missing
                    );
                }
                #endregion
                #region 문서 닫기
                doc.Close();
                #endregion
                logger.Info("Conversion success");
                returnValue.isSuccess = true;
                returnValue.Message = "Conversion was successful.";
                return returnValue;
            }
            catch (Exception e1)
            {
                logger.Error("======= Method: " + MethodBase.GetCurrentMethod().Name + " =======");
                logger.Error(new StackTrace(e1, true).ToString());
                logger.Error("Conversion failure: " + e1.Message);
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
