using DocConvert_Core.imageLib;
using DocConvert_Core.interfaces;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DocConvert_Core.WebCaptureLib
{
    public class WebCapture_Core
    {
        [STAThread]
        public static ReturnValue WebCapture(string Url, string outPath, ImageFormat imageFormat)
        {
            ReturnValue returnValue = new ReturnValue();
            try
            {


                returnValue.isSuccess = true;
                returnValue.Message = "WebCapture에 성공하였습니다.";
                returnValue.PageCount = 0;
            }
            catch (Exception e1)
            {
                returnValue.isSuccess = false;
                returnValue.Message = e1.Message;
                returnValue.PageCount = 0;
            }

            return returnValue;
        }
    }
}
