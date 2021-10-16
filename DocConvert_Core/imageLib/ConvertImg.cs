﻿using DocConvert_Core.interfaces;
using NLog;
using PdfiumViewer;
using System;
using System.Diagnostics;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Reflection;

namespace DocConvert_Core.imageLib
{
    public class ConvertImg
    {
        private static Logger logger = LogManager.GetLogger("DocConvert_Core_Log");

        // 참고문서: https://github.com/pvginkel/PdfiumViewer
        public static ReturnValue PDFtoJpeg(string SourcePDF, string outPath)
        {
            return PDFtoJpeg(SourcePDF, outPath, PdfRenderFlags.ForPrinting);
        }
        /// <summary>
        /// PDF에서 JPG로 변환하는 함수
        /// </summary>
        /// <param name="SourcePDF">대상 PDF 경로</param>
        /// <param name="outPath">내보낼 경로</param>
        /// <param name="quality">이미지 퀄리티</param>
        /// <returns>성공 여부(JSON)</returns>
        public static ReturnValue PDFtoJpeg(string SourcePDF, string outPath, PdfRenderFlags quality)
        {
            ReturnValue returnValue = new ReturnValue();
            try
            {
                logger.Info("==================== Start ====================");
                logger.Info("Method: " + MethodBase.GetCurrentMethod().Name + ", FilePath: " + SourcePDF + ", outPath: " + outPath);
                PdfDocument documentPage = PdfDocument.Load(SourcePDF);
                int pageCount = documentPage.PageCount;
                documentPage.Dispose();
                PdfDocument[] document = new PdfDocument[(pageCount / 100) + 1];
                int totalCnt = 0;
                int x = 0;
                for (int i = 0; i < (pageCount / 100) + 1; i++)
                {
                    document[i] = PdfDocument.Load(SourcePDF);

                    if (i == (pageCount / 100))
                    {
                        x = pageCount % 100;
                    }
                    else
                    {
                        x = 100;
                    }
                    
                    for (int j = i * 100; j < (i * 100) + x; j++)
                    {
                        int dpi = 300;
                        System.Drawing.Image imageInfo = document[i].Render(j, dpi, dpi, quality);
                        using (System.Drawing.Image image = document[i].Render(j, Convert.ToInt32(imageInfo.Width * 1.5), Convert.ToInt32(imageInfo.Height * 1.5), dpi, dpi, quality))
                        {
                            ImageCodecInfo encoder = ImageCodecInfo.GetImageEncoders()
                                .First(c => c.FormatID == ImageFormat.Jpeg.Guid);
                            EncoderParameters encParams = new EncoderParameters(1);
                            encParams.Param[0] = new EncoderParameter(Encoder.Quality, 100L);

                            image.Save(outPath + (j + 1) + ".jpg", encoder, encParams);
                            ++totalCnt;
                        }
                    }
                    document[i].Dispose();
                }
                if (pageCount == totalCnt)
                {
                    returnValue.PageCount = pageCount;
                    returnValue.isSuccess = true;
                    returnValue.Message = "Image conversion was successful.";
                    logger.Info("Image conversion was successful.");
                }
                else
                {
                    returnValue.PageCount = pageCount;
                    returnValue.isSuccess = false;
                    returnValue.Message = "Image conversion failed.";
                    logger.Error("Image conversion failed.");
                    new IOException("Image conversion failed.");
                }
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
                logger.Info("==================== End ====================");
            }
            return returnValue;
        }

        // 참고문서: https://github.com/pvginkel/PdfiumViewer
        public static ReturnValue PDFtoBmp(string SourcePDF, string outPath)
        {
            return PDFtoBmp(SourcePDF, outPath, PdfRenderFlags.ForPrinting);
        }
        /// <summary>
        /// PDF에서 BMP로 변환하는 함수
        /// </summary>
        /// <param name="SourcePDF">대상 PDF 경로</param>
        /// <param name="outPath">내보낼 경로</param>
        /// <param name="quality">이미지 퀄리티</param>
        /// <returns>성공 여부(JSON)</returns>
        public static ReturnValue PDFtoBmp(string SourcePDF, string outPath, PdfRenderFlags quality)
        {
            ReturnValue returnValue = new ReturnValue();
            try
            {
                logger.Info("==================== Start ====================");
                logger.Info("Method: " + MethodBase.GetCurrentMethod().Name + ", FilePath: " + SourcePDF + ", outPath: " + outPath);
                PdfDocument documentPage = PdfDocument.Load(SourcePDF);
                int pageCount = documentPage.PageCount;
                documentPage.Dispose();
                PdfDocument[] document = new PdfDocument[(pageCount / 100) + 1];
                int totalCnt = 0;
                int x = 0;
                for (int i = 0; i < (pageCount / 100) + 1; i++)
                {
                    document[i] = PdfDocument.Load(SourcePDF);

                    if (i == (pageCount / 100))
                    {
                        x = pageCount % 100;
                    }
                    else
                    {
                        x = 100;
                    }

                    for (int j = i * 100; j < (i * 100) + x; j++)
                    {
                        int dpi = 300;

                        System.Drawing.Image imageInfo = document[i].Render(j, dpi, dpi, quality);
                        using (System.Drawing.Image image = document[i].Render(j, Convert.ToInt32(imageInfo.Width * 1.5), Convert.ToInt32(imageInfo.Height * 1.5), dpi, dpi, quality))
                        {
                            ImageCodecInfo encoder = ImageCodecInfo.GetImageEncoders()
                                .First(c => c.FormatID == ImageFormat.Bmp.Guid);
                            EncoderParameters encParams = new EncoderParameters(1);
                            encParams.Param[0] = new EncoderParameter(Encoder.Quality, 100L);

                            image.Save(outPath + (j + 1) + ".bmp", encoder, encParams);
                            ++totalCnt;
                        }
                    }
                    document[i].Dispose();
                }
                if (pageCount == totalCnt)
                {
                    returnValue.PageCount = pageCount;
                    returnValue.isSuccess = true;
                    returnValue.Message = "Image conversion was successful.";
                    logger.Info("Image conversion was successful.");
                }
                else
                {
                    returnValue.PageCount = pageCount;
                    returnValue.isSuccess = false;
                    returnValue.Message = "Image conversion failed.";
                    logger.Error("Image conversion failed.");
                    new IOException("Image conversion failed.");
                }
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
                logger.Info("==================== End ====================");
            }
            return returnValue;
        }

        // 참고문서: https://github.com/pvginkel/PdfiumViewer
        public static ReturnValue PDFtoPng(string SourcePDF, string outPath)
        {
            return PDFtoPng(SourcePDF, outPath, PdfRenderFlags.ForPrinting);
        }
        /// <summary>
        /// PDF에서 PNG로 변환하는 함수
        /// </summary>
        /// <param name="SourcePDF">대상 PDF 경로</param>
        /// <param name="outPath">내보낼 경로</param>
        /// <param name="quality">이미지 퀄리티</param>
        /// <returns>성공 여부(JSON)</returns>
        public static ReturnValue PDFtoPng(string SourcePDF, string outPath, PdfRenderFlags quality)
        {

            ReturnValue returnValue = new ReturnValue();
            try
            {
                logger.Info("==================== Start ====================");
                logger.Info("Method: " + MethodBase.GetCurrentMethod().Name + ", FilePath: " + SourcePDF + ", outPath: " + outPath);
                PdfDocument documentPage = PdfDocument.Load(SourcePDF);
                int pageCount = documentPage.PageCount;
                documentPage.Dispose();
                PdfDocument[] document = new PdfDocument[(pageCount / 100) + 1];
                int totalCnt = 0;
                int x = 0;
                for (int i = 0; i < (pageCount / 100) + 1; i++)
                {
                    document[i] = PdfDocument.Load(SourcePDF);

                    if (i == (pageCount / 100))
                    {
                        x = pageCount % 100;
                    }
                    else
                    {
                        x = 100;
                    }

                    for (int j = i * 100; j < (i * 100) + x; j++)
                    {
                        int dpi = 300;

                        System.Drawing.Image imageInfo = document[i].Render(j, dpi, dpi, quality);
                        using (System.Drawing.Image image = document[i].Render(j, Convert.ToInt32(imageInfo.Width * 1.5), Convert.ToInt32(imageInfo.Height * 1.5), dpi, dpi, quality))
                        {
                            ImageCodecInfo encoder = ImageCodecInfo.GetImageEncoders()
                                .First(c => c.FormatID == ImageFormat.Png.Guid);
                            EncoderParameters encParams = new EncoderParameters(1);
                            encParams.Param[0] = new EncoderParameter(Encoder.Quality, 100L);

                            image.Save(outPath + (j + 1) + ".png", encoder, encParams);
                            ++totalCnt;
                        }
                    }
                    document[i].Dispose();
                }
                if (pageCount == totalCnt)
                {
                    returnValue.PageCount = pageCount;
                    returnValue.isSuccess = true;
                    returnValue.Message = "Image conversion was successful.";
                    logger.Info("Image conversion was successful.");
                }
                else
                {
                    returnValue.PageCount = pageCount;
                    returnValue.isSuccess = false;
                    returnValue.Message = "Image conversion failed.";
                    logger.Error("Image conversion failed.");
                    new IOException("Image conversion failed.");
                }
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
                logger.Info("==================== End ====================");
            }
            return returnValue;
        }

        /// <summary>
        /// PDF를 읽어서 페이지 카운트하는 로직
        /// </summary>
        /// <param name="SourcePDF">대상 PDF 경로</param>
        /// <returns>페이지 카운트</returns>
        public static int pdfPageCount(string SourcePDF)
        {
            try
            {
                using (PdfDocument document = PdfDocument.Load(SourcePDF))
                {
                    return document.PageCount;
                }
            }
            catch (Exception e1)
            {
                logger.Error("======= Method: " + MethodBase.GetCurrentMethod().Name + " =======");
                logger.Error(new StackTrace(e1, true).ToString());
                logger.Error("Image counting failure: " + e1.Message);
                logger.Error("================ End ================");
                return -1;
            }
        }
    }
}
