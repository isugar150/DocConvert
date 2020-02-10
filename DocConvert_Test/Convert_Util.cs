using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using DocConvert_Core.imageLib;
using DocConvert_Core.interfaces;
using DocConvert_Core.OfficeLib;

using Newtonsoft.Json.Linq;

using DocConvert_Core.HWPLib;
using Microsoft.Win32;
using System.Threading;

namespace DocConvert_Util
{
    public partial class Convert_Util : Form
    {
        private JObject Setting = new JObject();
        private bool HWPREGDLL = false;
        private bool APPVISIBLE = false;
        private bool RUNAFTER = false;
        private bool PAGINGNUM = false;
        public Convert_Util()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            #region JSON파일 파싱
            if (File.Exists(Application.StartupPath + @"\Settings.json"))
            {
                Setting = JObject.Parse(File.ReadAllText(Application.StartupPath + @"\Settings.json"));
                #region 변환창 보이기 초기데이터
                try
                {
                    if(bool.Parse(Setting["appvisible"].ToString()))
                    {
                        pictureBox4.Image = DocConvert_Util.Properties.Resources.switch_on;
                        APPVISIBLE = true;
                    }
                    else
                    {
                        pictureBox4.Image = DocConvert_Util.Properties.Resources.switch_off;
                        APPVISIBLE = false;
                    }
                }
                catch (Exception)
                {
                    pictureBox4.Image = DocConvert_Util.Properties.Resources.switch_on;
                    APPVISIBLE = true;
                    Setting["appvisible"] = APPVISIBLE;
                }
                #endregion
                #region 변환후 실행 초기 데이터
                try
                {
                    if (bool.Parse(Setting["runafter"].ToString()))
                    {
                        pictureBox6.Image = DocConvert_Util.Properties.Resources.switch_on;
                        RUNAFTER = true;
                    }
                    else
                    {
                        pictureBox6.Image = DocConvert_Util.Properties.Resources.switch_off;
                        RUNAFTER = false;
                    }
                }
                catch (Exception)
                {
                    pictureBox6.Image = DocConvert_Util.Properties.Resources.switch_on;
                    RUNAFTER = true;
                    Setting["runafter"] = RUNAFTER;
                }
                #endregion
                #region 페이징 번호
                try
                {
                    if (bool.Parse(Setting["pagingnum"].ToString()))
                    {
                        pictureBox8.Image = DocConvert_Util.Properties.Resources.switch_on;
                        PAGINGNUM = true;
                    }
                    else
                    {
                        pictureBox8.Image = DocConvert_Util.Properties.Resources.switch_off;
                        PAGINGNUM = false;
                    }
                }
                catch (Exception)
                {
                    pictureBox8.Image = DocConvert_Util.Properties.Resources.switch_on;
                    PAGINGNUM = true;
                    Setting["pagingnum"] = PAGINGNUM;
                }
                #endregion
            }
            else
            {
                tb2_appendText("[정보]   설정 파일이 존재하지 않습니다.");
            }
            #endregion
            #region 아래 한글 레지스트리 등록 확인
            RegistryKey regKey = Registry.CurrentUser.CreateSubKey(@"Software\HNC\HwpCtrl\Modules", RegistryKeyPermissionCheck.ReadWriteSubTree);
            try
            {
                if (File.Exists(regKey.GetValue("FilePathCheckerModuleExample").ToString())){
                    pictureBox1.Image = DocConvert_Util.Properties.Resources.switch_on;
                    HWPREGDLL = true;
                    tb2_appendText("[정보]   HWP DLL 경로: " + regKey.GetValue("FilePathCheckerModuleExample").ToString());
                }
                else
                {
                    tb2_appendText("[정보]   HWP DLL이 등록되어 있지 않습니다.");
                }
            }
            catch (Exception) {
                tb2_appendText("[정보]   HWP DLL이 등록되어 있지 않습니다.");
            }
            finally
            {
                regKey.Close();
            }
            #endregion
            comboBox1.SelectedIndex = 0;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.FileName = "";
            openFileDialog1.Filter = "지원하는 형식 (*.docx;*.doc;*.hwp;*.xlsx;*.xls;*.pptx;*.ppt;*.pdf)|*.docx;*.doc;*.hwp;*.xlsx;*.xls;*.pptx;*.ppt;*.ppt;*.pdf" +
                "|Word 형식 (*.docx;*.doc;*.hwp;)|*.docx;*.doc;*.hwp;" +
                "|Cell 형식 (*.xlsx;*.xls;)|*.xlsx;*.xls;" +
                "|PPT 형식 (*.pptx;*.ppt;)|*.pptx;*.ppt;" +
                "|PDF 형식 (*.pdf)|*.pdf" +
                "|All Files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;
            if(openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = "";
                for(int i = 0; i < openFileDialog1.FileNames.Length; i++)
                {
                    if ((i+1) == openFileDialog1.FileNames.Length)
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
            if (radioButton1.Checked)
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
                    if (Path.GetExtension(FileNames[i]).Equals(".hwp") && textBox3.Text.Length > 0)
                    {
                        tb2_appendText("비밀번호 해제후 다시시도해주세요.");
                        return;
                    }
                    ReturnValue status = null;
                    string passwd = null;
                    outPath = Path.GetDirectoryName(FileNames[i]) + @"\" + Path.GetFileNameWithoutExtension(FileNames[i]) + ".pdf";
                    tb2_appendText("[상태]   변환 요청 (" + (i + 1) + "/" + FileNames.Length + ")");
                    tb2_appendText("[정보]   소스 경로: " + FileNames[i]);
                    tb2_appendText("[정보]   출력 경로: " + outPath);
                    if (!textBox3.Text.Equals(""))
                        passwd = textBox3.Text;
                    if (Path.GetExtension(FileNames[i]).Equals(".docx") || Path.GetExtension(FileNames[i]).Equals(".doc"))
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
                    if(status != null) {
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
                            pdfToImgReturn = ConvertImg.PDFtoJpeg(outPath, imageOutput);
                        }
                        else if (comboBox1.SelectedIndex == 2)
                        {
                            pdfToImgReturn = ConvertImg.PDFtoPng(outPath, imageOutput);
                        }
                        else if (comboBox1.SelectedIndex == 3)
                        {
                            pdfToImgReturn = ConvertImg.PDFtoBmp(outPath, imageOutput);
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
            }
            else if (radioButton2.Checked)
            {
                #region Server 변환시

                #endregion
            }
        }

        #region 자잘한 이벤트
        
        // 로그 폴더 실행
        private void button4_Click(object sender, EventArgs e)
        {
            Process.Start(Application.StartupPath+@"\Log");
        }

        // IP입력창 키 이벤트
        private void ipAddressControl1_KeyUp(object sender, KeyEventArgs e)
        {
            Console.WriteLine("KeyUp: {0}", e.KeyValue);
        }
        
        // textBox2  문자열 추가
        private void tb2_appendText(string str)
        {
            textBox2.AppendText(System.DateTime.Now.ToString("yyyy/MM/dd hh:mm:ss.fff") + "   " + str + "\r\n");
        }
        
        // 한글 DLL 등록 버튼
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

        // 변환창 보이기 버튼
        private void pictureBox4_Click(object sender, EventArgs e)
        {
            #region 변환창 보이기 설정
            if (APPVISIBLE)
            {
                pictureBox4.Image = DocConvert_Util.Properties.Resources.switch_off;
                APPVISIBLE = false;
            }
            else
            {
                pictureBox4.Image = DocConvert_Util.Properties.Resources.switch_on;
                APPVISIBLE = true;
            }
            Setting["appvisible"] = APPVISIBLE;
            tb2_appendText("[정보]   변환창 보이기 설정: " + APPVISIBLE);
            try
            {
                File.WriteAllText(Application.StartupPath + @"\Settings.json", Setting.ToString());
            }
            catch (Exception e1)
            {
                tb2_appendText("[오류]   " + e1.Message);
            }
            #endregion
        }

        // 변환 후 실행 버튼
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
            Setting["runafter"] = RUNAFTER;
            tb2_appendText("[정보]   변환 후 실행 설정: " + RUNAFTER);
            try
            {
                File.WriteAllText(Application.StartupPath + @"\Settings.json", Setting.ToString());
            }
            catch (Exception e1)
            {
                tb2_appendText("[오류]   " + e1.Message);
            }
            #endregion
        }

        #endregion

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
            Setting["pagingnum"] = PAGINGNUM;
            tb2_appendText("[정보]   페이지 번호 추출: " + PAGINGNUM);
            try
            {
                File.WriteAllText(Application.StartupPath + @"\Settings.json", Setting.ToString());
            }
            catch (Exception e1)
            {
                tb2_appendText("[오류]   " + e1.Message);
            }
            #endregion
        }
    }
}
