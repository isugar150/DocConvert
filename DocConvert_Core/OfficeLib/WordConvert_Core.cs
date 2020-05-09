using DocConvert_Core.FileLib;
using DocConvert_Core.interfaces;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using NLog;
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.ServiceModel;
using Word = Microsoft.Office.Interop.Word;

namespace DocConvert_Core.OfficeLib
{
    public class WordConvert_Core
    {
        private static Logger logger = LogManager.GetLogger("DocConvert_Engine_Log");
        /// <summary>
        /// 워드파일을 PDF로 변환
        /// </summary>
        /// <param name="FilePath">소스파일</param>
        /// <param name="outPath">저장파일</param>
        /// <param name="docPassword">문서 비밀번호</param>
        /// <returns></returns>
        public static ReturnValue WordSaveAs(string FilePath, string outPath, string docPassword, bool pageCounting, bool appvisible)
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
            _Application word = null;
            try
            {
                if (appvisible)
                {
                    word = new Application
                    {
                        Visible = true,
                        DisplayAlerts = WdAlertLevel.wdAlertsNone
                    };
                }
                else
                {
                    word = new Application
                    {
                        Visible = false,
                        DisplayAlerts = WdAlertLevel.wdAlertsNone
                    };
                }
                // 매크로 실행 안되게 처리 (https://msdn.microsoft.com/en-us/library/microsoft.office.core.msoautomationsecurity.aspx?f=255&MSPPError=-2147217396)
                word.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;

                #region 열기 옵션 https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.documents.open?view=word-pia
                object ConfirmConversions = false;
                object ReadOnly = true;
                object AddToRecentFiles = false;
                object PasswordDocument = Type.Missing;
                if (docPassword != null)
                {
                    PasswordDocument = docPassword;
                }

                object PasswordTemplate = Type.Missing;
                object Revert = Type.Missing;
                object WritePasswordDocument = Type.Missing;
                object WritePasswordTemplate = Type.Missing;
                object Format = WdOpenFormat.wdOpenFormatAuto;
                /*object Format = WdOpenFormat.wdOpenFormatAllWord;
                if (Path.GetExtension(FilePath).Equals(".txt"))
                {
                    Format = WdOpenFormat.wdOpenFormatEncodedText;
                }
                else if (Path.GetExtension(FilePath).Equals(".html"))
                {
                    Format = WdOpenFormat.wdOpenFormatWebPages;
                }*/

                object Encoding = Type.Missing;
                object Visible = Type.Missing;
                object OpenAndRepair = Type.Missing;
                object DocumentDirection = Type.Missing;
                object NoEncodingDialog = Type.Missing;
                object XMLTransform = Type.Missing;
                #endregion
                #region 문서 열기
                _Document doc = word.Documents.Open(
                    FilePath,
                    ConfirmConversions,
                    ReadOnly,
                    AddToRecentFiles,
                    PasswordDocument,
                    PasswordTemplate,
                    Revert,
                    WritePasswordDocument,
                    WritePasswordTemplate,
                    Format,
                    Encoding,
                    Visible,
                    OpenAndRepair,
                    DocumentDirection,
                    NoEncodingDialog,
                    XMLTransform
                );

                doc.Activate();
                #endregion
                #region 페이지수 얻기
                if (pageCounting)
                {
                    returnValue.PageCount = doc.ComputeStatistics(WdStatistic.wdStatisticPages, -1);
                }
                #endregion
                #region 저장 옵션 https://docs.microsoft.com/ko-kr/dotnet/api/microsoft.office.tools.word.document.saveas2?view=vsto-2017
                // SaveAs2
                object FileFormat = WdSaveFormat.wdFormatPDF;
                object LockComments = Type.Missing;
                object Password = Type.Missing;
                object WritePassword = Type.Missing;
                object ReadOnlyRecommended = Type.Missing;
                object EmbedTrueTypeFonts = Type.Missing;
                object SaveNativePictureFormat = Type.Missing;
                object SaveFormsData = Type.Missing;
                object SaveAsAOCELetter = Type.Missing;
                object InsertLineBreaks = Type.Missing;
                object AllowSubstitutions = Type.Missing;
                object LineEnding = Type.Missing;
                object AddBiDiMarks = Type.Missing;
                object CompatibilityMode = Type.Missing;

                // ExportAsFixedFormat
                WdExportFormat ExportFormat = WdExportFormat.wdExportFormatPDF;
                bool OpenAfterExport = false;
                WdExportOptimizeFor OptimizeFor = WdExportOptimizeFor.wdExportOptimizeForPrint;
                WdExportRange Range = WdExportRange.wdExportAllDocument;
                Int32 From = 0;
                Int32 To = 0;
                WdExportItem Item = WdExportItem.wdExportDocumentContent;
                bool IncludeDocProps = false;
                bool KeepIRM = false;
                WdExportCreateBookmarks CreateBookmarks = WdExportCreateBookmarks.wdExportCreateNoBookmarks;
                bool DocStructureTags = false;
                bool BitmapMissingFonts = false;
                bool UseISO19005_1 = false;
                object FixedFormatExtClassPtr = Type.Missing;

                #endregion
                #region PDF저장
                try
                {
                    doc.SaveAs2(
                       outPath,
                       FileFormat,
                       LockComments,
                       Password,
                       AddToRecentFiles,
                       WritePassword,
                       ReadOnlyRecommended,
                       EmbedTrueTypeFonts,
                       SaveNativePictureFormat,
                       SaveFormsData,
                       SaveAsAOCELetter,
                       Encoding,
                       InsertLineBreaks,
                       AllowSubstitutions,
                       LineEnding,
                       AddBiDiMarks,
                       CompatibilityMode
                   );
                }
                catch (Exception)
                {
                    doc.ExportAsFixedFormat(
                        outPath,
                        ExportFormat,
                        OpenAfterExport,
                        OptimizeFor,
                        Range,
                        From,
                        To,
                        Item,
                        IncludeDocProps,
                        KeepIRM,
                        CreateBookmarks,
                        DocStructureTags,
                        BitmapMissingFonts,
                        UseISO19005_1,
                        FixedFormatExtClassPtr
                    );
                }
                #endregion
                #region 문서 닫기 옵션 https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.documents.close?view=word-pia
                object SaveChanges = WdSaveOptions.wdDoNotSaveChanges;
                object OriginalFormat = WdOriginalFormat.wdWordDocument;
                object RouteDocument = false;
                #endregion
                #region 문서 닫기
                doc.Close(SaveChanges, OriginalFormat, RouteDocument);
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
                word.Quit();
                Marshal.ReleaseComObject(word);
                word = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                #endregion
                logger.Info("==================== End ====================");
            }
        }
    }
}
