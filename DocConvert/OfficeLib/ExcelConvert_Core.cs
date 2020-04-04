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
        public static ReturnValue ExcelSaveAs(string FilePath, string outPath, string docPassword, bool pageCounting, bool appvisible)
        {
            ReturnValue returnValue = new ReturnValue();
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

                #region 열기 옵션 https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.workbooks.open?view=excel-pia
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
                #region 페이지수 얻기
                if (pageCounting)
                {
                    int sheetCount = 0, pageCount = 0;
                    Sheets sheets = null;
                    Worksheet sheet = null;
                    PageSetup pageSetup = null;
                    Pages pages = null;
                    try
                    {
                        sheets = doc.Worksheets;
                        sheetCount = sheets.Count;

                        for (int index = 1; index <= sheetCount; index++)
                        {
                            sheet = (Excel.Worksheet)sheets[index];
                            sheet.Activate();

                            pageSetup = sheet.PageSetup;

                            pageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;
                            pageSetup.PaperSize = Excel.XlPaperSize.xlPaperA4;

                            pages = pageSetup.Pages;
                            pageCount += pages.Count;
                        }

                        returnValue.PageCount = pageCount;
                    }
                    catch (Exception e1)
                    {
                        returnValue.PageCount = -1;
                        logger.Error("페이지 카운트 가져오는중 오류발생");
                        logger.Error("오류내용: " + e1.Message);
                    }
                }
                #endregion
                #region 저장 옵션 https://docs.microsoft.com/ko-kr/dotnet/api/microsoft.office.tools.excel.workbook.exportasfixedformat?view=vsto-2017
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
                #region  문서 닫기 옵션 https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._workbook.close?view=excel-pia
                object SaveChanges = false;
                object Filename = Type.Missing;
                object RouteWorkbook = false;
                #endregion
                #region 문서 닫기
                doc.Close(SaveChanges, Filename, RouteWorkbook);
                #endregion
                logger.Info("변환 성공");
                returnValue.isSuccess = true;
                returnValue.Message = "변환에 성공하였습니다.";
                return returnValue;
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
                excel.Quit();
                Marshal.ReleaseComObject(excel);
                excel = null;
                excel = new Application { };
                excel.Quit();
                Marshal.ReleaseComObject(excel);
                excel = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                #endregion
                logger.Info("==================== End ====================");
            }
        }
    }
}
