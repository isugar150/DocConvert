using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.Word;
using NLog;
using Word = Microsoft.Office.Interop.Word;

namespace DocConvert.OfficeLib
{
    class WordConvert_Core
    {
        private static Logger logger = LogManager.GetLogger("DocConvert_Log");
        /// <summary>
        /// 워드파일을 PDF로 변환
        /// </summary>
        /// <param name="FilePath">소스파일</param>
        /// <param name="outPath">저장파일</param>
        /// <param name="docPassword">문서 비밀번호</param>
        /// <param name="format">저장 형식</param>
        /// <returns></returns>
        public static bool WordSaveAs(String FilePath, String outPath, String docPassword)
        {
            logger.Info("==================== Start ====================");
            logger.Info("Method: WordSaveAs, FilePath: " + FilePath + ", outPath: " + outPath + ", docPassword: " + docPassword);
            try
            {
                _Application word = new Word.Application
                {
                    Visible = false
                };

                #region 열기 옵션
                object UpdateLinks = false;
                object ReadOnly = true;
                object Format = Type.Missing;
                object Password = Type.Missing;
                if (docPassword != null)
                    Password = docPassword;
                object WriteResPassword = Type.Missing;
                object IgnoreReadOnlyRecommended = true;
                object Origin = Type.Missing;
                object Delimiter = Type.Missing;
                object Editable = false;
                object Notify = false;
                object Converter = Type.Missing;
                object AddToMru = Type.Missing;
                object Local = false;
                object CorruptLoad = Type.Missing;
                #endregion
                #region 문서 열기
                _Document doc = word.Documents.Open(
                    FilePath,
                    UpdateLinks,
                    ReadOnly,
                    Format,
                    Password,
                    WriteResPassword,
                    IgnoreReadOnlyRecommended,
                    Origin,
                    Delimiter,
                    Editable,
                    Notify,
                    Converter,
                    AddToMru,
                    Local,
                    CorruptLoad
                );

                doc.Activate();
                #endregion
                #region 저장 옵션
                object FileFormat = WdSaveFormat.wdFormatPDF;
                object LockComments = Type.Missing;
                object AddToRecentFiles = Type.Missing;
                object WritePassword = Type.Missing;
                object ReadOnlyRecommended = Type.Missing;
                object EmbedTrueTypeFonts = Type.Missing;
                object SaveNativePictureFormat = Type.Missing;
                object SaveFormsData = Type.Missing;
                object SaveAsAOCELetter = Type.Missing;
                object Encoding = Type.Missing;
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
                #region 문서 종료
                object SaveChanges = Type.Missing;
                object OriginalFormat = Type.Missing;
                object RouteDocument = Type.Missing;

                word.Quit(SaveChanges, OriginalFormat, RouteDocument);
                #endregion
                logger.Info("변환 성공");
                return true;
            }
            catch(Exception e1)
            {
                logger.Error(e1.Message);
                return false;
            }
            finally
            {
                logger.Info("==================== End ====================");
            }
        }
    }
}
