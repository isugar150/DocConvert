using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using DocConvert_Core.interfaces;
using DocConvert_Core.FileLib;
using Microsoft.Office.Interop.Word;
using NLog;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Core;

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
        public static ReturnValue WordSaveAs(String FilePath, String outPath, String docPassword, bool pageCounting, bool appvisible)
        {
            ReturnValue returnValue = new ReturnValue();
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
            _Application word = null;
            try
            {
                if (appvisible)
                {
                    word = new Word.Application
                    {
                        Visible = true,
                        DisplayAlerts = WdAlertLevel.wdAlertsNone
                    };
                }
                else
                {
                    word = new Word.Application
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
                    PasswordDocument = docPassword;
                object PasswordTemplate = Type.Missing;
                object Revert = Type.Missing;
                object WritePasswordDocument = Type.Missing;
                object WritePasswordTemplate = Type.Missing;
                object Format = Type.Missing;
                object Encoding = Type.Missing;
                object Visible = false;
                object OpenAndRepair = Type.Missing;
                object DocumentDirection = Type.Missing;
                object NoEncodingDialog = Type.Missing;
                object XMLTransform = Type.Missing;
                #endregion
                #region 문서 열기
                _Document doc = word.Documents.Open(
                    FilePath,
                    ReadOnly,
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
                    returnValue.PageCount = doc.ComputeStatistics(WdStatistic.wdStatisticPages, -1);
                #endregion
                #region 저장 옵션 https://docs.microsoft.com/ko-kr/dotnet/api/microsoft.office.tools.word.document.saveas2?view=vsto-2017
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
                #endregion
                #region PDF저장
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
                #endregion
                #region 문서 닫기 옵션 https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.documents.close?view=word-pia
                object SaveChanges = WdSaveOptions.wdDoNotSaveChanges;
                object OriginalFormat = WdOriginalFormat.wdWordDocument;
                object RouteDocument = true;
                #endregion
                #region 문서 닫기
                doc.Close(SaveChanges, OriginalFormat, RouteDocument);
                #endregion
                logger.Info("변환 성공");
                returnValue.isSuccess = true;
                returnValue.Message = "변환에 성공하였습니다.";
                return returnValue;
            }
            catch(Exception e1)
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
                word.Quit();
                Marshal.ReleaseComObject(word);
                word = null;
                #endregion
                logger.Info("==================== End ====================");
            }
        }
    }
}
