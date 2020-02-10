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

namespace DocConvert_Core.imageLib
{
    public class ConvertImg
    {
        // 참고문서: https://github.com/pvginkel/PdfiumViewer
        public static ReturnValue PDFtoJpeg(string SourcePDF, string outPath)
        {
            ReturnValue returnValue = new ReturnValue();
            try
            {
                using (var document = PdfDocument.Load(SourcePDF))
                {
                    var pageCount = document.PageCount;
                    for (int i = 0; i < pageCount; i++)
                    {
                        Debug.WriteLine(outPath + (i + 1) + ".jpg");
                        var dpi = 300;

                        using (var image = document.Render(i, dpi, dpi, PdfRenderFlags.CorrectFromDpi))
                        {
                            var encoder = ImageCodecInfo.GetImageEncoders()
                                .First(c => c.FormatID == ImageFormat.Jpeg.Guid);
                            var encParams = new EncoderParameters(1);
                            encParams.Param[0] = new EncoderParameter(
                                System.Drawing.Imaging.Encoder.Quality, 100L);

                            image.Save(outPath + (i + 1) + ".jpg", encoder, encParams);

                            returnValue.PageCount = pageCount;
                            returnValue.isSuccess = true;
                            returnValue.Message = "이미지 변환에 성공하였습니다.";
                        }
                    }
                }
            }
            catch (Exception e1)
            {
                returnValue.PageCount = -1;
                returnValue.isSuccess = false;
                returnValue.Message = e1.Message;
            }
            return returnValue;
        }

        // 참고문서: https://github.com/pvginkel/PdfiumViewer
        public static ReturnValue PDFtoBmp(string SourcePDF, string outPath)
        {
            ReturnValue returnValue = new ReturnValue();
            try
            {
                using (var document = PdfDocument.Load(SourcePDF))
                {
                    var pageCount = document.PageCount;
                    for (int i = 0; i < pageCount; i++)
                    {
                        Debug.WriteLine(outPath + (i + 1) + ".bmp");
                        var dpi = 300;

                        using (var image = document.Render(i, dpi, dpi, PdfRenderFlags.CorrectFromDpi))
                        {
                            var encoder = ImageCodecInfo.GetImageEncoders()
                                .First(c => c.FormatID == ImageFormat.Bmp.Guid);
                            var encParams = new EncoderParameters(1);
                            encParams.Param[0] = new EncoderParameter(
                                System.Drawing.Imaging.Encoder.Quality, 100L);

                            image.Save(outPath + (i + 1) + ".bmp", encoder, encParams);

                            returnValue.PageCount = pageCount;
                            returnValue.isSuccess = true;
                            returnValue.Message = "이미지 변환에 성공하였습니다.";
                        }
                    }
                }
            }
            catch (Exception e1)
            {
                returnValue.PageCount = -1;
                returnValue.isSuccess = false;
                returnValue.Message = e1.Message;
            }
            return returnValue;
        }

        // 참고문서: https://github.com/pvginkel/PdfiumViewer
        public static ReturnValue PDFtoPng(string SourcePDF, string outPath)
        {
            ReturnValue returnValue = new ReturnValue();
            try
            {
                using (var document = PdfDocument.Load(SourcePDF))
                {
                    var pageCount = document.PageCount;
                    for (int i = 0; i < pageCount; i++)
                    {
                        Debug.WriteLine(outPath + (i + 1) + ".png");
                        var dpi = 300;

                        using (var image = document.Render(i, dpi, dpi, PdfRenderFlags.CorrectFromDpi))
                        {
                            var encoder = ImageCodecInfo.GetImageEncoders()
                                .First(c => c.FormatID == ImageFormat.Png.Guid);
                            var encParams = new EncoderParameters(1);
                            encParams.Param[0] = new EncoderParameter(
                                System.Drawing.Imaging.Encoder.Quality, 100L);

                            image.Save(outPath + (i + 1) + ".png", encoder, encParams);

                            returnValue.PageCount = pageCount;
                            returnValue.isSuccess = true;
                            returnValue.Message = "이미지 변환에 성공하였습니다.";
                        }
                    }
                }
            }
            catch (Exception e1)
            {
                returnValue.PageCount = -1;
                returnValue.isSuccess = false;
                returnValue.Message = e1.Message;
            }
            return returnValue;
        }
    }
}
