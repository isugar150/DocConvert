using PdfiumViewer;
using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DocConvert_Core.interfaces;
using System.Diagnostics;
using NLog;
using System.Reflection;

namespace DocConvert_Core.imageLib
{
    public class ConvertImg
    {
        private static Logger logger = LogManager.GetLogger("DocConvert_Engine_Log");

        // 참고문서: https://github.com/pvginkel/PdfiumViewer
        public static ReturnValue PDFtoJpeg(string SourcePDF, string outPath)
        {
            return PDFtoJpeg(SourcePDF, outPath, PdfRenderFlags.ForPrinting);
        }
        public static ReturnValue PDFtoJpeg(string SourcePDF, string outPath, PdfRenderFlags quality)
        {
            ReturnValue returnValue = new ReturnValue();
            try
            {
                logger.Info("==================== Start ====================");
                logger.Info("Method: " + MethodBase.GetCurrentMethod().Name + ", FilePath: " + SourcePDF + ", outPath: " + outPath);
                using (var document = PdfDocument.Load(SourcePDF))
                {
                    var pageCount = document.PageCount;
                    int totalCnt = 0;
                    for (int i = 0; i < pageCount; i++)
                    {
                        var dpi = 300;

                        using (var image = document.Render(i, dpi, dpi, quality))
                        {
                            var encoder = ImageCodecInfo.GetImageEncoders()
                                .First(c => c.FormatID == ImageFormat.Jpeg.Guid);
                            var encParams = new EncoderParameters(1);
                            encParams.Param[0] = new EncoderParameter(
                                System.Drawing.Imaging.Encoder.Quality, 100L);

                            image.Save(outPath + (i + 1) + ".jpg", encoder, encParams);
                            ++totalCnt;
                        }
                    }
                    if(pageCount == totalCnt)
                    {
                        returnValue.PageCount = pageCount;
                        returnValue.isSuccess = true;
                        returnValue.Message = "이미지 변환에 성공하였습니다.";
                        logger.Info("이미지 변환 성공!");
                    }
                    else
                    {
                        returnValue.PageCount = pageCount;
                        returnValue.isSuccess = false;
                        returnValue.Message = "이미지 변환에 실패하였습니다.";
                        logger.Error("이미지 변환 실패!");
                        new IOException("이미지 변환에 실패하였습니다.");
                    }
                }
            }
            catch (Exception e1)
            {
                throw e1;
            }
            finally
            {
                logger.Info("==================== End ====================");
            }
            return returnValue;
        }

        // 참고문서: https://github.com/pvginkel/PdfiumViewer
        public static ReturnValue PDFtoBmp(string SourcePDF, string outPath)
        {
            return PDFtoBmp(SourcePDF, outPath, PdfRenderFlags.ForPrinting);
        }
        public static ReturnValue PDFtoBmp(string SourcePDF, string outPath, PdfRenderFlags quality)
        {
            ReturnValue returnValue = new ReturnValue();
            try
            {
                logger.Info("==================== Start ====================");
                logger.Info("Method: " + MethodBase.GetCurrentMethod().Name + ", FilePath: " + SourcePDF + ", outPath: " + outPath);
                using (var document = PdfDocument.Load(SourcePDF))
                {
                    var pageCount = document.PageCount;
                    int totalCnt = 0;
                    for (int i = 0; i < pageCount; i++)
                    {
                        Debug.WriteLine(outPath + (i + 1) + ".bmp");
                        var dpi = 300;

                        using (var image = document.Render(i, dpi, dpi, quality))
                        {
                            var encoder = ImageCodecInfo.GetImageEncoders()
                                .First(c => c.FormatID == ImageFormat.Bmp.Guid);
                            var encParams = new EncoderParameters(1);
                            encParams.Param[0] = new EncoderParameter(
                                System.Drawing.Imaging.Encoder.Quality, 100L);

                            image.Save(outPath + (i + 1) + ".bmp", encoder, encParams);
                            ++totalCnt;
                        }
                    }
                    if (pageCount == totalCnt)
                    {
                        returnValue.PageCount = pageCount;
                        returnValue.isSuccess = true;
                        returnValue.Message = "이미지 변환에 성공하였습니다.";
                        logger.Info("이미지 변환 성공!");
                    }
                    else
                    {
                        returnValue.PageCount = pageCount;
                        returnValue.isSuccess = false;
                        returnValue.Message = "이미지 변환에 실패하였습니다.";
                        logger.Error("이미지 변환 실패!");
                        new IOException("이미지 변환에 실패하였습니다.");
                    }
                }
            }
            catch (Exception e1)
            {
                throw e1;
            }
            finally
            {
                logger.Info("==================== End ====================");
            }
            return returnValue;
        }

        // 참고문서: https://github.com/pvginkel/PdfiumViewer
        public static ReturnValue PDFtoPng(string SourcePDF, string outPath)
        {
            return PDFtoPng(SourcePDF, outPath, PdfRenderFlags.ForPrinting);
        }
        public static ReturnValue PDFtoPng(string SourcePDF, string outPath, PdfRenderFlags quality)
        {
            ReturnValue returnValue = new ReturnValue();
            try
            {
                logger.Info("==================== Start ====================");
                logger.Info("Method: " + MethodBase.GetCurrentMethod().Name + ", FilePath: " + SourcePDF + ", outPath: " + outPath);
                using (var document = PdfDocument.Load(SourcePDF))
                {
                    var pageCount = document.PageCount;
                    int totalCnt = 0;
                    for (int i = 0; i < pageCount; i++)
                    {
                        Debug.WriteLine(outPath + (i + 1) + ".png");
                        var dpi = 300;

                        using (var image = document.Render(i, dpi, dpi, quality))
                        {
                            var encoder = ImageCodecInfo.GetImageEncoders()
                                .First(c => c.FormatID == ImageFormat.Png.Guid);
                            var encParams = new EncoderParameters(1);
                            encParams.Param[0] = new EncoderParameter(
                                System.Drawing.Imaging.Encoder.Quality, 100L);

                            image.Save(outPath + (i + 1) + ".png", encoder, encParams);
                            ++totalCnt;
                        }
                    }
                    if (pageCount == totalCnt)
                    {
                        returnValue.PageCount = pageCount;
                        returnValue.isSuccess = true;
                        returnValue.Message = "이미지 변환에 성공하였습니다.";
                        logger.Info("이미지 변환 성공!");
                    }
                    else
                    {
                        returnValue.PageCount = pageCount;
                        returnValue.isSuccess = false;
                        returnValue.Message = "이미지 변환에 실패하였습니다.";
                        logger.Error("이미지 변환 실패!");
                        new IOException("이미지 변환에 실패하였습니다.");
                    }
                }
            }
            catch (Exception e1)
            {
                throw e1;
            }
            finally
            {
                logger.Info("==================== End ====================");
            }
            return returnValue;
        }

        public static int pdfPageCount(string SourcePDF)
        {
            try
            {
                using (var document = PdfDocument.Load(SourcePDF))
                {
                    return document.PageCount;
                }
            }
            catch (Exception e1)
            {
                logger.Info("이미지 카운팅 실패!");
                logger.Info("Source image: " + SourcePDF);
                logger.Info(e1.Message);
                return -1;
            }
        }
    }
}
