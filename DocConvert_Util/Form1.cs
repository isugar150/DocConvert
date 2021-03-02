using DocConvert_Core.HWPLib;
using DocConvert_Core.imageLib;
using DocConvert_Core.interfaces;
using DocConvert_Core.OfficeLib;
using DocConvert_Core.ZipLib;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DocConvert_Util
{
    public partial class Form1 : Form
    {
        string openDialogFilter = "지원하는 형식 (*.docx;*.doc;*.hwp;*.xlsx;*.xls;*.csv;*.pptx;*.ppt;*.pdf;*.txt;*.html;)|*.docx;*.doc;*.hwp;*.xlsx;*.xls;*.csv;*.pptx;*.ppt;*.ppt;*.pdf;*.txt;*.html;" +
                "|Word 형식 (*.docx;*.doc;*.hwp;*.txt;*.html;)|*.docx;*.doc;*.hwp;*.txt;*.html;" +
                "|Cell 형식 (*.xlsx;*.xls;*.csv;)|*.xlsx;*.xls*.csv;;" +
                "|PPT 형식 (*.pptx;*.ppt;)|*.pptx;*.ppt;" +
                "|PDF 형식 (*.pdf)|*.pdf" +
                "|All Files (*.*)|*.*";

        Socket sock = null;

        public Form1()
        {
            InitializeComponent();
        }

        // 폼이 로드할때 실행
        private void Form1_Load(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = 0;
            comboBox2.SelectedIndex = 0;
            textBox8.Enabled = false;
        }

        // Local-선택 버튼
        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = openDialogFilter;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
            }
        }

        // Local-실행 버튼
        private void button2_Click(object sender, EventArgs e)
        {
            ReturnValue status = new ReturnValue();
            string targetFile = textBox1.Text;
            string newPdfFile = Path.GetDirectoryName(targetFile) + @"\" + Path.GetFileNameWithoutExtension(targetFile) + ".pdf";
            string newZipFile = Path.GetDirectoryName(targetFile) + @"\" + Path.GetFileNameWithoutExtension(targetFile) + ".zip";
            string docPassword = textBox3.Text;
            bool PAGINGNUM = false;
            bool APPVISIBLE = false;
            int convertImg = comboBox1.SelectedIndex;
            
            outputText("[로컬 변환 시작]");
            outputText("파일 경로: " + targetFile);
            outputText("이미지 변환: " + comboBox1.Text);

            if (Path.GetExtension(targetFile).Equals(".docx") || Path.GetExtension(targetFile).Equals(".doc") || Path.GetExtension(targetFile).Equals(".txt") || Path.GetExtension(targetFile).Equals(".html"))
            {
                status = WordConvert_Core.WordSaveAs(targetFile, newPdfFile, docPassword, PAGINGNUM, APPVISIBLE);
            }
            else if (Path.GetExtension(targetFile).Equals(".xlsx") || Path.GetExtension(targetFile).Equals(".xls") || Path.GetExtension(targetFile).Equals(".csv"))
            {
                status = ExcelConvert_Core.ExcelSaveAs(targetFile, newPdfFile, docPassword, PAGINGNUM, APPVISIBLE);
            }
            else if (Path.GetExtension(targetFile).Equals(".pptx") || Path.GetExtension(targetFile).Equals(".ppt"))
            {
                status = PowerPointConvert_Core.PowerPointSaveAs(targetFile, newPdfFile, docPassword, PAGINGNUM, APPVISIBLE);
            }
            else if (Path.GetExtension(targetFile).Equals(".hwp"))
            {
                Thread HWPConvert = new Thread(() =>
                {
                    status = HWPConvert_Core.HwpSaveAs(targetFile, newPdfFile, PAGINGNUM);
                });
                HWPConvert.SetApartmentState(ApartmentState.STA);
                HWPConvert.Start();
                HWPConvert.Join();

            }
            else if (Path.GetExtension(targetFile).Equals(".pdf"))
            {
                if (convertImg == 0)
                {
                    status.isSuccess = true;
                    status.Message = "PDF파일을 PDF파일로 변환할 수 없습니다.";
                    status.PageCount = ConvertImg.pdfPageCount(targetFile);
                }
            }
            outputText("isSuccess: " + status.isSuccess);
            outputText("Message: " + status.Message);
            outputText("PageCount: " + ConvertImg.pdfPageCount(newPdfFile));

            // 이미지 변환
            if (convertImg != 0)
            {
                DirectoryInfo imageOutPath = new DirectoryInfo(Path.GetDirectoryName(targetFile) + @"\" + Path.GetFileNameWithoutExtension(targetFile) + "\\");
                if (!imageOutPath.Exists)
                    imageOutPath.Create();
                if (convertImg == 1)
                {
                    status = ConvertImg.PDFtoJpeg(newPdfFile, imageOutPath.FullName, PdfiumViewer.PdfRenderFlags.ForPrinting);
                }
                else if (convertImg == 2)
                {
                    status = ConvertImg.PDFtoPng(newPdfFile, imageOutPath.FullName, PdfiumViewer.PdfRenderFlags.ForPrinting);
                }
                else if (convertImg == 3)
                {
                    status = ConvertImg.PDFtoBmp(newPdfFile, imageOutPath.FullName, PdfiumViewer.PdfRenderFlags.ForPrinting);
                }

                // 압축
                if (status.isSuccess)
                    ZipLib.CreateZipFile(Directory.GetFiles(imageOutPath.FullName), newZipFile);
                outputText("isSuccess: " + status.isSuccess);
                outputText("Message: " + status.Message);
                outputText("PageCount: " + ConvertImg.pdfPageCount(newPdfFile));
            }
        }

        // Remote-선택 버튼
        private void button3_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = openDialogFilter;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox2.Text = openFileDialog1.FileName;
            }
        }

        // Remote-실행 버튼
        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                outputText("[서버 변환 시작]");
                FileInfo srcFile = new FileInfo(textBox2.Text);
                DirectoryInfo tmpDir = new DirectoryInfo(textBox9.Text);
                FileInfo targetFile = new FileInfo(tmpDir + @"\" + Path.GetFileName(srcFile.FullName));

                if (!srcFile.Exists)
                {
                    outputText("소스 파일 경로가 없습니다.");
                    return;
                }
                if (!tmpDir.Exists)
                {
                    outputText("tmp폴더가 존재하지 않습니다.");
                    return;
                } 

                srcFile.CopyTo(targetFile.FullName, true);

                sock = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
                IPEndPoint ep = new IPEndPoint(IPAddress.Parse(textBox5.Text), int.Parse(textBox6.Text));
                byte[] receiverBuff = new byte[1024];

                sock.Connect(ep);

                JObject sendJson = new JObject();
                sendJson["ClientKey"] = textBox7.Text;
                sendJson["Method"] = "DocConvert";
                sendJson["FileName"] = Path.GetFileName(srcFile.FullName);
                sendJson["ConvertImg"] = comboBox2.SelectedIndex;
                sendJson["DocPassword"] = textBox4.Text;
                sendJson["DRM_UseYn"] = checkBox1.Checked ? "Y":"n";
                sendJson["DRM_Type"] = textBox8.Text;

                outputText("Send Data\r\n" + sendJson.ToString());

                byte[] sendData = Encoding.UTF8.GetBytes(sendJson.ToString() + "\r\n");
                int bytesSent = sock.Send(sendData, sendData.Length, SocketFlags.None);

                int welcome = sock.Receive(receiverBuff);
                int bytesRec = sock.Receive(receiverBuff);
                string data = Encoding.UTF8.GetString(receiverBuff, 0, bytesRec);
                outputText("Recive Data\r\n" + data);

                sock.Shutdown(SocketShutdown.Receive);
                //sock.Close();
            } catch(Exception e1)
            {
                outputText(e1.Message);
                outputText(e1.StackTrace);
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                textBox8.Enabled = true;
            }
            else
            {
                textBox8.Enabled = false;
            }
        }

        private void outputText(string text)
        {
            output.AppendText(text + "\r\n");
        }
    }
}
