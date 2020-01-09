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

using DocConvert_Core.OfficeLib;

using Newtonsoft.Json.Linq;

using DocConvert_Core.HWPLib;
using Microsoft.Win32;

namespace DocConvert_Util
{
    public partial class Convert_Util : Form
    {
        private JObject Setting = new JObject();
        private bool HWPREGDLL = false;
        private bool APPVISIBLE = false;
        private bool RUNAFTER = false;
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
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.FileName = "";
            openFileDialog1.Filter = "지원하는 형식 (*.docx;*.doc;*.hwp;*.xlsx;*.xls;*.pptx;*.ppt)|*.docx;*.doc;*.hwp;*.xlsx;*.xls;*.pptx;*.ppt" +
                "|Word 형식 (*.docx;*.doc;*.hwp;)|*.docx;*.doc;*.hwp;" +
                "|Cell 형식 (*.xlsx;*.xls;)|*.xlsx;*.xls;" +
                "|PPT 형식 (*.pptx;*.ppt;)|*.pptx;*.ppt;" +
                "|All Files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;

            if(openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                #region Local 변환시
                if (Path.GetExtension(textBox1.Text).Equals(".hwp") && textBox3.Text.Length > 0)
                {
                    MessageBox.Show("비밀번호 해제후 다시시도해주세요.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                bool status = false;
                string passwd = null;
                string outPath = Path.GetDirectoryName(textBox1.Text) + @"\" + Path.GetFileNameWithoutExtension(textBox1.Text) + ".pdf";
                tb2_appendText("[상태]   변환 요청");
                tb2_appendText("[정보]   소스 경로: " + textBox1.Text);
                tb2_appendText("[정보]   출력 경로: " + Path.GetDirectoryName(textBox1.Text) + @"\" + Path.GetFileNameWithoutExtension(textBox1.Text) + ".pdf");
                if (!textBox3.Text.Equals(""))
                    passwd = textBox3.Text;
                if (Path.GetExtension(textBox1.Text).Equals(".docx") || Path.GetExtension(textBox1.Text).Equals(".doc"))
                {
                    status = WordConvert_Core.WordSaveAs(textBox1.Text, outPath, passwd, APPVISIBLE);
                    if (status)
                        tb2_appendText("[상태]   변환 성공");
                    else
                        tb2_appendText("[상태]   변환 실패");
                }
                else if (Path.GetExtension(textBox1.Text).Equals(".xlsx") || Path.GetExtension(textBox1.Text).Equals(".xls"))
                {
                    status = ExcelConvert_Core.ExcelSaveAs(textBox1.Text, outPath, passwd, APPVISIBLE);
                    if (status)
                        tb2_appendText("[상태]   변환 성공");
                    else
                        tb2_appendText("[상태]   변환 실패");
                }
                else if (Path.GetExtension(textBox1.Text).Equals(".pptx") || Path.GetExtension(textBox1.Text).Equals(".ppt"))
                {
                    status = PowerPointConvert_Core.PowerPointSaveAs(textBox1.Text, outPath, passwd, APPVISIBLE);
                    if (status)
                        tb2_appendText("[상태]   변환 성공");
                    else
                        tb2_appendText("[상태]   변환 실패");
                }
                else if (Path.GetExtension(textBox1.Text).Equals(".hwp"))
                {
                    status = HWPConvert_Core.HwpSaveAs(textBox1.Text, outPath);
                    if (status)
                        tb2_appendText("[상태]   변환 성공");
                    else
                        tb2_appendText("[상태]   변환 실패");
                }
                else
                {
                    tb2_appendText("[상태]   지원포맷 아님. 파싱한 확장자: " + Path.GetExtension(textBox1.Text));
                }
                #endregion
                #region 변환 후 실행
                if (RUNAFTER && status)
                {
                    Process.Start(Path.GetDirectoryName(textBox1.Text) + @"\" + Path.GetFileNameWithoutExtension(textBox1.Text) + ".pdf");
                }
                #endregion
            }
            else if (radioButton2.Checked)
            {
                #region Server 변환시

                #endregion
            }
        }

        #region 자잘한 이벤트

        private void button4_Click(object sender, EventArgs e)
        {
            Process.Start(Application.StartupPath+@"\Log");
        }

        private void ipAddressControl1_KeyUp(object sender, KeyEventArgs e)
        {
            Console.WriteLine("KeyUp: {0}", e.KeyValue);
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            groupBox2.Enabled = false;
            ipAddressControl1.Enabled = true;
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            groupBox2.Enabled = true;
            ipAddressControl1.Enabled = false;
        }

        private void tb2_appendText(string str)
        {
            textBox2.AppendText(System.DateTime.Now.ToString("yyyy/MM/dd hh:mm:ss.fff") + "   " + str + "\r\n");
        }

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
    }
}
