using DocConvert_Core.HWPLib;
using DocConvert_Core.imageLib;
using DocConvert_Core.IniLib;
using DocConvert_Core.interfaces;
using DocConvert_Core.OfficeLib;
using FluentFTP;
using Microsoft.Win32;
using Newtonsoft.Json.Linq;
using PdfiumViewer;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WebSocketSharp;

namespace DocConvert_Util
{
    public partial class Convert_Util : Form
    {
        private bool HWPREGDLL = false;
        private bool APPVISIBLE = false;
        private bool RUNAFTER = false;
        private bool PAGINGNUM = false;
        private DateTime timeTaken;
        private PdfRenderFlags quality = PdfRenderFlags.ForPrinting;
        public static iniProperties IniProperties = new iniProperties();
        public Convert_Util()
        {
            InitializeComponent();
            CheckForIllegalCrossThreadCalls = false;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            checkBox1.Checked = true;

            #region Parse INI
            IniFile pairs = new IniFile();
            while (true)
            {
                try
                {
                    if (new FileInfo("./DocConvert_Util.ini").Exists)
                    {
                        pairs.Load("./DocConvert_Util.ini");
                        IniProperties.TargetIP = pairs["DC Util"]["targetIP"].ToString();
                        IniProperties.ftpUser = pairs["DC Util"]["ftpUser"].ToString();
                        IniProperties.ftpPwd = pairs["DC Util"]["ftpPwd"].ToString();
                        IniProperties.socketPort = int.Parse(pairs["DC Util"]["socketPort"].ToString());
                        IniProperties.filePort = int.Parse(pairs["DC Util"]["filePort"].ToString());
                        IniProperties.isFTPS = pairs["DC Util"]["isFTPS"].ToString().Equals("Y");
                        IniProperties.appvisible = pairs["DC Util"]["appvisible"].ToString().Equals("Y");
                        IniProperties.runafter = pairs["DC Util"]["runafter"].ToString().Equals("Y");
                        IniProperties.pagingnum = pairs["DC Util"]["pagingnum"].ToString().Equals("Y");

                        textBox4.Text = IniProperties.TargetIP;
                        textBox5.Text = IniProperties.socketPort.ToString();
                        textBox6.Text = IniProperties.filePort.ToString();
                        textBox8.Text = IniProperties.ftpUser;
                        textBox7.Text = IniProperties.ftpPwd;

                        checkBox2.Checked = IniProperties.isFTPS;

                        if (IniProperties.appvisible)
                        {
                            pictureBox4.Image = Properties.Resources.switch_on;
                            APPVISIBLE = true;
                        }
                        if (IniProperties.runafter)
                        {
                            pictureBox6.Image = Properties.Resources.switch_on;
                            RUNAFTER = true;
                        }
                        if (IniProperties.pagingnum)
                        {
                            pictureBox8.Image = Properties.Resources.switch_on;
                            PAGINGNUM = true;
                        }

                        break;
                    }
                    else
                    {
                        Setting.createSetting();
                        MessageBox.Show("설정 파일을 생성하였습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                catch (Exception)
                {
                    if (MessageBox.Show("설정 파일이 손상되었습니다. 초기화하시겠습니까?", "알림", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                    {
                        Setting.createSetting();
                        MessageBox.Show("설정 파일을 초기화하였습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            #endregion
            #region 아래 한글 레지스트리 등록 확인
            RegistryKey regKey = Registry.CurrentUser.CreateSubKey(@"Software\HNC\HwpCtrl\Modules", RegistryKeyPermissionCheck.ReadWriteSubTree);
            try
            {
                if (File.Exists(regKey.GetValue("FilePathCheckerModuleExample").ToString()))
                {
                    pictureBox1.Image = DocConvert_Util.Properties.Resources.switch_on;
                    HWPREGDLL = true;
                    tb2_appendText("[정보]   HWP DLL 경로: " + regKey.GetValue("FilePathCheckerModuleExample").ToString());
                }
                else
                {
                    tb2_appendText("[정보]   HWP DLL이 등록되어 있지 않습니다.");
                }
            }
            catch (Exception)
            {
                tb2_appendText("[정보]   HWP DLL이 등록되어 있지 않습니다.");
            }
            finally
            {
                regKey.Close();
            }
            #endregion
            comboBox1.SelectedIndex = 0;
            comboBox2.SelectedIndex = 0;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.FileName = "";
            openFileDialog1.Filter = "지원하는 형식 (*.docx;*.doc;*.hwp;*.xlsx;*.xls;*.pptx;*.ppt;*.pdf;*.txt;*.html;)|*.docx;*.doc;*.hwp;*.xlsx;*.xls;*.pptx;*.ppt;*.ppt;*.pdf;*.txt;*.html;" +
                "|Word 형식 (*.docx;*.doc;*.hwp;*.txt;*.html;)|*.docx;*.doc;*.hwp;*.txt;*.html;" +
                "|Cell 형식 (*.xlsx;*.xls;)|*.xlsx;*.xls;" +
                "|PPT 형식 (*.pptx;*.ppt;)|*.pptx;*.ppt;" +
                "|PDF 형식 (*.pdf)|*.pdf" +
                "|All Files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = "";
                for (int i = 0; i < openFileDialog1.FileNames.Length; i++)
                {
                    if ((i + 1) == openFileDialog1.FileNames.Length)
                    {
                        textBox1.AppendText(openFileDialog1.FileNames[i]);

                    }
                    else
                    {
                        textBox1.AppendText(openFileDialog1.FileNames[i] + "|");
                    }
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                string serverIP = textBox4.Text;
                int serverPORT = int.Parse(textBox5.Text);
                int filePORT = int.Parse(textBox6.Text);
                string ftpUser = textBox8.Text;
                string ftpPwd = textBox7.Text;
                timeTaken = DateTime.Now;
                groupBox1.Enabled = false;
                groupBox2.Enabled = false;
                if (comboBox2.SelectedIndex == 0)
                {
                    // 유효성 검사
                    if (textBox1.Text.Equals("") || !new FileInfo(textBox1.Text).Exists) //파일이 없으면
                    {
                        tb2_appendText("파일이 존재하지 않습니다.");
                        groupBox1.Enabled = true;
                        groupBox2.Enabled = true;
                        return;
                    }
                    if (!checkBox1.Checked)
                    {
                        #region Local 변환시
                        String[] FileNames = textBox1.Text.Split('|');
                        String outPath = null;
                        short SuccessCount = 0;
                        for (int i = 0; i < FileNames.Length; i++)
                        {
                            if (!File.Exists(FileNames[i]))
                            {
                                tb2_appendText(FileNames[i] + "해당 경로에 파일이 존재하지 않습니다.");
                            }
                            ReturnValue status = null;
                            string passwd = null;
                            outPath = Path.GetDirectoryName(FileNames[i]) + @"\" + Path.GetFileNameWithoutExtension(FileNames[i]) + ".pdf";
                            tb2_appendText("[상태]   변환 요청 (" + (i + 1) + "/" + FileNames.Length + ")");
                            tb2_appendText("[정보]   소스 경로: " + FileNames[i]);
                            tb2_appendText("[정보]   출력 경로: " + outPath);
                            if (Path.GetExtension(FileNames[i]).Equals(".docx") || Path.GetExtension(FileNames[i]).Equals(".doc") || Path.GetExtension(FileNames[i]).Equals(".txt") || Path.GetExtension(FileNames[i]).Equals(".html"))
                            {
                                status = WordConvert_Core.WordSaveAs(FileNames[i], outPath, passwd, PAGINGNUM, APPVISIBLE);
                            }
                            else if (Path.GetExtension(FileNames[i]).Equals(".xlsx") || Path.GetExtension(FileNames[i]).Equals(".xls"))
                            {
                                status = ExcelConvert_Core.ExcelSaveAs(FileNames[i], outPath, passwd, PAGINGNUM, APPVISIBLE);
                            }
                            else if (Path.GetExtension(FileNames[i]).Equals(".pptx") || Path.GetExtension(FileNames[i]).Equals(".ppt"))
                            {
                                status = PowerPointConvert_Core.PowerPointSaveAs(FileNames[i], outPath, passwd, PAGINGNUM, APPVISIBLE);
                            }
                            else if (Path.GetExtension(FileNames[i]).Equals(".hwp"))
                            {
                                status = HWPConvert_Core.HwpSaveAs(FileNames[i], outPath, PAGINGNUM);
                            }
                            else if (Path.GetExtension(FileNames[i]).Equals(".pdf"))
                            {
                                ReturnValue pdfreturnValue = new ReturnValue();
                                pdfreturnValue.isSuccess = true;
                                pdfreturnValue.Message = "PDF파일은 변환할 필요가 없습니다.";
                                pdfreturnValue.PageCount = ConvertImg.pdfPageCount(FileNames[i]);
                                status = pdfreturnValue;
                            }
                            else 
                            {
                                tb2_appendText("[상태]   지원포맷 아님. 파싱한 확장자: " + Path.GetExtension(FileNames[i]));
                            }
                            if (status != null)
                            {
                                if (PAGINGNUM)
                                    tb2_appendText("[정보]   페이지 수: " + Convert.ToString(status.PageCount));

                                if (status.isSuccess)
                                    tb2_appendText("[상태]   " + status.Message);
                                else
                                    tb2_appendText("[오류]   " + status.Message);
                                if (status.isSuccess)
                                    SuccessCount += 1;
                            }
                            // PDF To Image
                            if (comboBox1.SelectedIndex != 0)
                            {
                                ReturnValue pdfToImgReturn = null;
                                String imageOutput = Path.GetDirectoryName(outPath) + "\\" + Path.GetFileNameWithoutExtension(outPath) + "\\";
                                if (!new DirectoryInfo(imageOutput).Exists)
                                    new DirectoryInfo(imageOutput).Create();
                                if (comboBox1.SelectedIndex == 1)
                                {
                                    pdfToImgReturn = ConvertImg.PDFtoJpeg(outPath, imageOutput, quality);
                                }
                                else if (comboBox1.SelectedIndex == 2)
                                {
                                    pdfToImgReturn = ConvertImg.PDFtoPng(outPath, imageOutput, quality);
                                }
                                else if (comboBox1.SelectedIndex == 3)
                                {
                                    pdfToImgReturn = ConvertImg.PDFtoBmp(outPath, imageOutput, quality);
                                }
                                if (pdfToImgReturn.isSuccess)
                                {
                                    tb2_appendText("[상태]   " + pdfToImgReturn.Message);
                                    tb2_appendText("[상태]   이미지로 내보낸 페이지수: " + pdfToImgReturn.PageCount);
                                }
                                else
                                {
                                    tb2_appendText("[오류]   " + pdfToImgReturn.Message);
                                }
                            }
                            #region 변환 후 실행
                            if (RUNAFTER && status.isSuccess)
                            {
                                Process.Start(Path.GetDirectoryName(FileNames[i]) + @"\" + Path.GetFileNameWithoutExtension(FileNames[i]) + ".pdf");
                            }
                            #endregion
                        }
                        if (FileNames.Length == SuccessCount)
                            tb2_appendText("[상태]   모든 파일을 변환완료 하였습니다.");
                        else
                            tb2_appendText("[상태]   일부 파일 변환을 실패하였습니다.");
                        #endregion
                        groupBox1.Enabled = true;
                        groupBox2.Enabled = true;
                        TimeSpan curTime = DateTime.Now - timeTaken;
                        tb2_appendText(string.Format("작업 소요시간: {0}", curTime.ToString()));
                    }
                    else
                    {
                        #region Server 변환시
                        try
                        {
                            using (var ftpClient = new FtpClient())
                            {
                                ftpClient.Host = serverIP;
                                ftpClient.Port = filePORT;
                                ftpClient.Credentials = new NetworkCredential(ftpUser, ftpPwd);
                                if (checkBox2.Checked)
                                {
                                    ftpClient.EncryptionMode = FtpEncryptionMode.Implicit;
                                    ftpClient.ValidateAnyCertificate = true;
                                }
                                tb2_appendText(serverIP + ":" + filePORT + "에 연결을 시도합니다.");
                                ftpClient.Connect();
                                tb2_appendText("File Server  " + serverIP + ":" + filePORT + "에 연결하였습니다.");

                                ftpClient.UploadFile(textBox1.Text, "tmp/" + Path.GetFileName(textBox1.Text), FtpRemoteExists.Overwrite, true);
                                tb2_appendText("서버에 파일을 업로드하였습니다.");
                                ftpClient.Disconnect();
                                ftpClient.Dispose();
                            }
                        }
                        catch (Exception e1)
                        {
                            tb2_appendText(e1.Message);
                            groupBox1.Enabled = true;
                            groupBox2.Enabled = true;
                            return;
                        }
                        // 아이피주소 유효성 검사
                        try
                        {
                            IPAddress.Parse(serverIP);
                        }
                        catch (Exception) { tb2_appendText("유효한 아이피 주소를 입력하세요."); groupBox1.Enabled = true; groupBox2.Enabled = true; return; }
                        // 포트번호 유효성 검사
                        if (serverPORT < IPEndPoint.MinPort || serverPORT > IPEndPoint.MaxPort)
                        {
                            tb2_appendText("유효한 포트번호를 입력하세요.");
                            groupBox1.Enabled = true;
                            groupBox2.Enabled = true;
                            return;
                        }
                        JObject requestMsg = new JObject();
                        requestMsg["KEY"] = "B29D00A3 - F825 - 4EB7 - 93C1 - A77F5E31A7C2";
                        requestMsg["Method"] = "DocConvert";
                        requestMsg["FileName"] = new FileInfo(textBox1.Text).Name;
                        requestMsg["ConvertIMG"] = comboBox1.SelectedIndex;
                        if(comboBox2.SelectedIndex == 0)
                            requestMsg["useCompression"] = checkBox3.Checked;

                        SendData(requestMsg.ToString());
                        #endregion
                    }
                }
                else if (comboBox2.SelectedIndex == 1)
                {
                    if (serverPORT < IPEndPoint.MinPort || serverPORT > IPEndPoint.MaxPort)
                    {
                        tb2_appendText("유효한 포트번호를 입력하세요.");
                        groupBox1.Enabled = true;
                        groupBox2.Enabled = true;
                        return;
                    }
                    JObject requestMsg = new JObject();
                    requestMsg["KEY"] = "B29D00A3 - F825 - 4EB7 - 93C1 - A77F5E31A7C2";
                    requestMsg["Method"] = "WebCapture";
                    requestMsg["URL"] = textBox9.Text;
                    SendData(requestMsg.ToString());
                }
            }
            catch (Exception e1)
            {
                groupBox1.Enabled = true;
                groupBox2.Enabled = true;
                tb2_appendText("변환중 오류발생");
                tb2_appendText("======= Method: " + MethodBase.GetCurrentMethod().Name + " =======");
                tb2_appendText(new StackTrace(e1, true).ToString());
                tb2_appendText("변환 실패: " + e1.Message);
                tb2_appendText("================ End ================");
            }
        }
        #region Socket Server Method
        /// <summary>
        /// 엡소켓연결 메소드
        /// </summary>
        /// <param name="Message">서버에 던질 메시지</param>
        async void SendData(string Message)
        {
            try
            {
                string targetURL = "ws://" + textBox4.Text + ":" + textBox5.Text;
                tb2_appendText(targetURL + "에 연결을 시도합니다.");
                WebSocket websocketClient = new WebSocket(url: targetURL);
                websocketClient.Connect();

                tb2_appendText(targetURL + "에 보낸 데이터\r\n" + Message);
                websocketClient.Send(Encoding.UTF8.GetBytes(Message));

                websocketClient.OnMessage += (sender, e) =>
                {
                    tb2_appendText("서버와 연결을 해제하였습니다.");
                    websocketClient.Close();
                    tb2_appendText(targetURL + "에서 받은 데이터\r\n" + e.Data);
                    JObject responseData = JObject.Parse(e.Data);

                    string isSuccess = responseData["isSuccess"].ToString();
                    string url = null;
                    try
                    {
                        url = responseData["URL"].ToString();
                    }
                    catch { }
                    string convertImgCnt = null;
                    try
                    {
                        convertImgCnt = responseData["convertImgCnt"].ToString();
                    }
                    catch { }
                    if (responseData["Method"].ToString().Equals("DocConvert"))
                    {
                    #region DocConvert
                    string imgType = comboBox1.Text.Replace("<", "").Replace(">", "");
                        string serverIP = textBox4.Text;
                        int filePORT = int.Parse(textBox6.Text);
                        string ftpUser = textBox8.Text;
                        string ftpPwd = textBox7.Text;
                        using (var ftpClient = new FtpClient())
                        {
                            ftpClient.Host = serverIP;
                            ftpClient.Port = filePORT;
                            ftpClient.Credentials = new NetworkCredential(ftpUser, ftpPwd);
                            if (checkBox2.Checked)
                            {
                                ftpClient.EncryptionMode = FtpEncryptionMode.Implicit;
                                ftpClient.ValidateAnyCertificate = true;
                            }
                            ftpClient.Connect();
                            if (isSuccess.Equals("True"))
                            {
                                string outPath = Path.GetDirectoryName(textBox1.Text);
                                string outFileName = Path.GetFileNameWithoutExtension(textBox1.Text);
                                if (comboBox1.SelectedIndex != 0)
                                {
                                    if (responseData["useCompression"].ToString().Equals("True"))
                                    {
                                        ftpClient.DownloadFile(outPath + @"\" + outFileName + ".zip", responseData["URL"].ToString() + @"\" + outFileName + ".zip", FtpLocalExists.Overwrite, FtpVerify.None);
                                        tb2_appendText(outPath + @"\" + outFileName + ".zip 파일 다운로드 완료");
                                    }
                                    else
                                    {
                                        for (int i = 0; i < int.Parse(convertImgCnt); i++)
                                        {
                                            ftpClient.DownloadFile(outPath + @"\" + outFileName + @"\" + (i + 1) + "." + imgType, url + "/" + outFileName + "/" + (i + 1) + "." + imgType, FtpLocalExists.Overwrite, FtpVerify.None);
                                            tb2_appendText(outPath + @"\" + outFileName + @"\" + (i + 1) + "." + imgType + " 파일 다운로드 완료");
                                        }
                                    }
                                }
                                if (comboBox1.SelectedIndex == 0)
                                {
                                    ftpClient.DownloadFile(outPath + @"\" + outFileName + ".pdf", url + "/" + outFileName + ".pdf", FtpLocalExists.Overwrite, FtpVerify.None);
                                    tb2_appendText(outPath + @"\" + outFileName + ".pdf" + " 파일 다운로드 완료");
                                }
                            }
                            ftpClient.Disconnect();
                            ftpClient.Dispose();
                        }
                    #endregion
                }
                    else if (responseData["Method"].ToString().Equals("WebCapture"))
                    {
                    #region WebCapture
                    string serverIP = textBox4.Text;
                        int filePORT = int.Parse(textBox6.Text);
                        string ftpUser = textBox8.Text;
                        string ftpPwd = textBox7.Text;
                        using (var ftpClient = new FtpClient())
                        {
                            ftpClient.Host = serverIP;
                            ftpClient.Port = filePORT;
                            ftpClient.Credentials = new NetworkCredential(ftpUser, ftpPwd);
                            if (checkBox2.Checked)
                            {
                                ftpClient.EncryptionMode = FtpEncryptionMode.Implicit;
                                ftpClient.ValidateAnyCertificate = true;
                            }
                            if (isSuccess.Equals("True"))
                            {
                                ftpClient.DownloadFile(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\" + new FileInfo(url).Name, url, FtpLocalExists.Overwrite, FtpVerify.None);
                                tb2_appendText("이미지를 바탕화면에 저장하였습니다.");
                            }
                            ftpClient.Disconnect();
                            ftpClient.Dispose();
                        }
                    #endregion
                }
                    groupBox1.Enabled = true;
                    groupBox2.Enabled = true;
                    TimeSpan curTime = DateTime.Now - timeTaken;
                    tb2_appendText(string.Format("작업 소요시간: {0}", curTime.ToString()));

                    IniProperties.TargetIP = textBox4.Text;
                    IniProperties.ftpUser = textBox8.Text;
                    IniProperties.ftpPwd = textBox7.Text;
                    IniProperties.socketPort = int.Parse(textBox5.Text);
                    IniProperties.filePort = int.Parse(textBox6.Text);
                    IniProperties.isFTPS = checkBox2.Checked;
                    Setting.updateSetting(IniProperties);
                };
            } catch(Exception e1)
            {
                groupBox1.Enabled = true;
                groupBox2.Enabled = true;
                tb2_appendText(e1.Message);
            }
        }

        #endregion

        #region 컴포넌트 이벤트

        /// <summary>
        /// 로그폴더 실행
        /// </summary>
        private void button4_Click(object sender, EventArgs e)
        {
            Process.Start(Application.StartupPath + @"\Log");
        }

        /// <summary>
        /// IP입력창 키 이벤트
        /// </summary>
        private void ipAddressControl1_KeyUp(object sender, KeyEventArgs e)
        {
            Debug.WriteLine("KeyUp: {0}", e.KeyValue);
        }

        /// <summary>
        /// textBox2  문자열 추가
        /// </summary>
        /// <param name="str">Append할 텍스트</param>
        private void tb2_appendText(string str)
        {
            textBox2.AppendText(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff") + "   " + str + "\r\n");
            Debug.WriteLine(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff") + "   " + str + "\r\n");
        }

        /// <summary>
        /// 한글 DLL 등록 버튼
        /// </summary>
        private void pictureBox1_Click(object sender, EventArgs e)
        {
            #region 한글 DLL 레지스트리 관리
            if (HWPREGDLL)
            {
                RegistryKey regKey = Registry.CurrentUser.CreateSubKey(@"Software\HNC\HwpCtrl\Modules", RegistryKeyPermissionCheck.ReadWriteSubTree);
                try
                {
                    Registry.CurrentUser.DeleteSubKey(@"Software\HNC\HwpCtrl\Modules");
                    tb2_appendText("[정보]   HWP DLL을 레지스트리에서 등록 해제 했습니다.");
                }
                catch (Exception e1)
                {
                    tb2_appendText("[오류]   레지스트리 등록 해제중 다음 오류 발생: " + e1.Message);
                    return;
                }
                finally
                {
                    regKey.Close();
                }

                pictureBox1.Image = DocConvert_Util.Properties.Resources.switch_off;
                HWPREGDLL = false;
            }
            else
            {
                if (File.Exists(Application.StartupPath + @"\FilePathCheckerModuleExample.dll"))
                {
                    RegistryKey regKey = Registry.CurrentUser.CreateSubKey(@"Software\HNC\HwpCtrl\Modules", RegistryKeyPermissionCheck.ReadWriteSubTree);
                    try
                    {
                        regKey.SetValue("FilePathCheckerModuleExample", Application.StartupPath + @"\FilePathCheckerModuleExample.dll", RegistryValueKind.String);
                        tb2_appendText("[정보]   HWP DLL을 레지스트리에 등록하였습니다.");

                        pictureBox1.Image = DocConvert_Util.Properties.Resources.switch_on;
                        HWPREGDLL = true;
                    }
                    catch (Exception e1)
                    {
                        tb2_appendText("[오류]   레지스트리 등록 해제중 다음 오류 발생: " + e1.Message);
                        return;
                    }
                    finally
                    {
                        regKey.Close();
                    }
                }
                else
                {
                    tb2_appendText("[오류]   레지스트리 등록에 필요한 파일이 존재하지 않습니다.\r\n확인후 다시시도하세요.");
                }
            }
            #endregion
        }

        /// <summary>
        /// 변환창 보이기 버튼
        /// </summary>
        private void pictureBox4_Click(object sender, EventArgs e)
        {
            #region 변환창 보이기 설정
            if (APPVISIBLE)
            {
                pictureBox4.Image = Properties.Resources.switch_off;
                APPVISIBLE = false;
            }
            else
            {
                pictureBox4.Image = Properties.Resources.switch_on;
                APPVISIBLE = true;
            }
            IniProperties.appvisible = APPVISIBLE;
            tb2_appendText("[정보]   변환창 보이기 설정: " + APPVISIBLE);
            try
            {
                Setting.updateSetting(IniProperties);
            }
            catch (Exception e1)
            {
                tb2_appendText("[오류]   " + e1.Message);
            }
            #endregion
        }

        /// <summary>
        /// 변환 후 실행 버튼
        /// </summary>
        private void pictureBox6_Click(object sender, EventArgs e)
        {
            #region 파일 변환 후 실행
            if (RUNAFTER)
            {
                pictureBox6.Image = DocConvert_Util.Properties.Resources.switch_off;
                RUNAFTER = false;
            }
            else
            {
                pictureBox6.Image = DocConvert_Util.Properties.Resources.switch_on;
                RUNAFTER = true;
            }
            IniProperties.runafter = RUNAFTER;
            tb2_appendText("[정보]   변환 후 실행 설정: " + RUNAFTER);
            try
            {
                Setting.updateSetting(IniProperties);
            }
            catch (Exception e1)
            {
                tb2_appendText("[오류]   " + e1.Message);
            }
            #endregion
        }

        /// <summary>
        /// 페이지 추출 버튼 (서버는 사용안함)
        /// </summary>
        private void pictureBox8_Click(object sender, EventArgs e)
        {
            #region 페이지 번호 추출
            if (PAGINGNUM)
            {
                pictureBox8.Image = DocConvert_Util.Properties.Resources.switch_off;
                PAGINGNUM = false;
            }
            else
            {
                pictureBox8.Image = DocConvert_Util.Properties.Resources.switch_on;
                PAGINGNUM = true;
            }
            IniProperties.pagingnum = PAGINGNUM;
            tb2_appendText("[정보]   페이지 번호 추출: " + PAGINGNUM);
            try
            {
                Setting.updateSetting(IniProperties);
            }
            catch (Exception e1)
            {
                tb2_appendText("[오류]   " + e1.Message);
            }
            #endregion
        }

        /// <summary>
        /// 서버 IP 체크박스 클릭 시
        /// </summary>
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                textBox4.Enabled = true;
                textBox5.Enabled = true;
                textBox6.Enabled = true;
                textBox7.Enabled = true;
                textBox8.Enabled = true;
                panel2.Enabled = false;
                panel3.Enabled = false;
                panel4.Enabled = false;
                panel5.Enabled = true;
                checkBox2.Enabled = true;
            }
            else
            {
                comboBox2.SelectedIndex = 0;
                textBox4.Enabled = false;
                textBox5.Enabled = false;
                textBox6.Enabled = false;
                textBox7.Enabled = false;
                textBox8.Enabled = false;
                panel2.Enabled = true;
                panel3.Enabled = true;
                panel4.Enabled = true;
                panel5.Enabled = false;
                checkBox2.Enabled = false;
            }
        }

        /// <summary>
        /// 변환 메소드 콤보박스
        /// </summary>
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox2.SelectedIndex == 0)
            {
                panel6.Location = new Point(0, 78);
                panel6.Visible = true;
                panel7.Visible = false;
                checkBox3.Enabled = true;
            }
            else if (comboBox2.SelectedIndex == 1)
            {
                panel7.Location = new Point(0, 109);
                panel7.Visible = true;
                panel6.Visible = false;
                checkBox3.Enabled = false;
            }
        }

        /// <summary>
        /// 변환할 이미지 콤보박스
        /// </summary>
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(comboBox1.SelectedIndex != 0 && comboBox2.SelectedIndex == 0)
            {
                checkBox3.Enabled = true;
            }
            else
            {
                checkBox3.Enabled = false;
            }
        }

        /// <summary>
        /// 해당 창을 닫을때 실행하는 이벤트
        /// </summary>
        private void Convert_Util_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                IniProperties.TargetIP = textBox4.Text;
                IniProperties.ftpUser = textBox8.Text;
                IniProperties.ftpPwd = textBox7.Text;
                IniProperties.socketPort = int.Parse(textBox5.Text);
                IniProperties.filePort = int.Parse(textBox6.Text);
                IniProperties.isFTPS = checkBox2.Checked;
                Setting.updateSetting(IniProperties);
            }
            catch (Exception e1)
            {
                tb2_appendText("[오류]   " + e1.Message);
            }
        }

        #endregion
    }
}