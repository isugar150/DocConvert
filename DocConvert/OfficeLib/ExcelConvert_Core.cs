﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using DocConvert_Core.FileLib;
using Microsoft.Office.Interop.Excel;
using NLog;
using Excel = Microsoft.Office.Interop.Excel;

namespace DocConvert_Core.OfficeLib
{
    public class ExcelConvert_Core
    {
        private static Logger logger = LogManager.GetLogger("DocConvert_Engine_Log");
        /// <summary>
        /// 워드파일을 PDF로 변환
        /// </summary>
        /// <param name="FilePath">소스파일</param>
        /// <param name="outPath">저장파일</param>
        /// <param name="docPassword">문서 비밀번호</param>
        /// <returns></returns>
        public static bool ExcelSaveAs(String FilePath, String outPath, String docPassword, bool appvisible)
        {
            logger.Info("==================== Start ====================");
            logger.Info("Method: " + MethodBase.GetCurrentMethod().Name + ", FilePath: " + FilePath + ", outPath: " + outPath + ", docPassword: " + docPassword);
            #region File Unlock
            try
            {
                LockFile.UnLock_File(FilePath);
                logger.Info("파일 언락 성공!");
            } catch(Exception e1)
            {
                logger.Info("파일 언락 실패! 자세한내용 로그 참고");
                logger.Error(e1.Message);
            }
            #endregion
            _Application excel = null;
            try
            {
                if (appvisible)
                {
                    excel = new Excel.Application
                    {
                        Visible = true,
                        DisplayAlerts = false
                    };
                }
                else
                {
                    excel = new Excel.Application
                    {
                        Visible = false,
                        DisplayAlerts = false
                    };
                }

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
                object AddToMru = false;
                object Local = false;
                object CorruptLoad = Type.Missing;
                #endregion
                #region 문서 열기
                _Workbook doc = excel.Workbooks.Open(
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
                XlFixedFormatType FileFormat = XlFixedFormatType.xlTypePDF;
                XlFixedFormatQuality Quality = XlFixedFormatQuality.xlQualityMinimum;
                object IncludeDocProperties = Type.Missing;
                object IgnorePrintAreas = false;
                object From = Type.Missing;
                object To = Type.Missing;
                object OpenAfterPublish = false;
                object FixedFormatExtClassPtr = Type.Missing;
                #endregion
                #region PDF저장
                doc.ExportAsFixedFormat(
                    FileFormat,
                    outPath,
                    Quality,
                    IncludeDocProperties,
                    IgnorePrintAreas,
                    From,
                    To,
                    OpenAfterPublish,
                    FixedFormatExtClassPtr
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
                excel.Quit();
                Marshal.ReleaseComObject(excel);
                excel = null;
                #endregion
                logger.Info("==================== End ====================");
            }
        }
    }
}
