using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.PowerPoint;
using NetOffice.OfficeApi.Enums;
using NLog;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using MsoTriState = Microsoft.Office.Core.MsoTriState;

namespace DocConvert.OfficeLib
{
    class PowerPointConvert_Core
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
        public static bool PowerPointSaveAs(String FilePath, String outPath, String docPassword)
        {
            logger.Info("==================== Start ====================");
            logger.Info("Method: PowerPointSaveAs, FilePath: " + FilePath + ", outPath: " + outPath + ", docPassword: " + docPassword);
            try
            {
                _Application powerpoint = new Application();
                Presentations multiPresentations = powerpoint.Presentations;

                #region 열기 옵션
                MsoTriState ReadOnly = MsoTriState.msoTrue;
                MsoTriState Untitled = MsoTriState.msoFalse;
                MsoTriState WithWindow = MsoTriState.msoFalse;
                #endregion
                #region 문서 열기
                Presentation doc = multiPresentations.Open(
                    FilePath,
                    ReadOnly,
                    Untitled,
                    WithWindow
                );
                #endregion
                #region 저장 옵션
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
                #region 문서 종료
                powerpoint.Quit();
                #endregion
                logger.Info("변환 성공");
                return true;
            }
            catch (Exception e1)
            {
                logger.Error("변환 실패: " + e1.Message);
                return false;
            }
            finally
            {
                logger.Info("==================== End ====================");
            }
        }
    }
}
