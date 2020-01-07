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

using DocConvert.OfficeLib;

using DocConvert.HWPLib;
using Microsoft.Win32;

namespace DocConvert
{
    public partial class Convert_Util : Form
    {
        private bool HWPREGDLL = false;
        public Convert_Util()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            textBox1.Text = @"C:\Users\Jm\Desktop\test.pptx";
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
            if (Path.GetExtension(textBox1.Text).Equals(".hwp") && textBox3.Text.Length > 0)
            {
                MessageBox.Show("비밀번호 해제후 다시시도해주세요.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            label3.Text = "상태: 변환중";
            bool status = false;
            string passwd = null;
            string outPath = Path.GetDirectoryName(textBox1.Text) + @"\" + Path.GetFileNameWithoutExtension(textBox1.Text) + ".pdf";
            if (!textBox3.Text.Equals(""))
                passwd = textBox3.Text;
            if (Path.GetExtension(textBox1.Text).Equals(".docx") || Path.GetExtension(textBox1.Text).Equals(".doc"))
            {
                status = WordConvert_Core.WordSaveAs(textBox1.Text, outPath, passwd);
                if (status)
                    label3.Text = "상태: 변환 성공";
                else
                    label3.Text = "상태: 변환 실패";
            }
            else if (Path.GetExtension(textBox1.Text).Equals(".xlsx") || Path.GetExtension(textBox1.Text).Equals(".xls"))
            {
                status = ExcelConvert_Core.ExcelSaveAs(textBox1.Text, outPath, passwd);
                if (status)
                    label3.Text = "상태: 변환 성공";
                else
                    label3.Text = "상태: 변환 실패";
            }
            else if (Path.GetExtension(textBox1.Text).Equals(".pptx") || Path.GetExtension(textBox1.Text).Equals(".ppt"))
            {
                status = PowerPointConvert_Core.PowerPointSaveAs(textBox1.Text, outPath, passwd);
                if (status)
                    label3.Text = "상태: 변환 성공";
                else
                    label3.Text = "상태: 변환 실패";
            }
            else if (Path.GetExtension(textBox1.Text).Equals(".hwp"))
            {
                status = HWPConvert_Core.HwpSaveAs(textBox1.Text, outPath);
                if (status)
                    label3.Text = "상태: 변환 성공";
                else
                    label3.Text = "상태: 변환 실패";
            }
            else
            {
                label3.Text = "상태: 지원포맷 아님";
            }
        }

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
            ipAddressControl1.Enabled = true;
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            ipAddressControl1.Enabled = false;
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            if (HWPREGDLL)
            {
                /*RegistryKey regKey = Registry.CurrentUser.CreateSubKey(@"Software\HNC\HwpCtrl\Modules", RegistryKeyPermissionCheck.ReadWriteSubTree);
                try
                {
                    regKey.DeleteSubKey("FilePathCheckerModuleExample", true);
                }
                catch(Exception e1)
                {
                    MessageBox.Show(e1.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                finally
                {
                    regKey.Close();
                }

                pictureBox1.Image = DocConvert_Util.Properties.Resources.switch_off;
                HWPREGDLL = false;*/
            }
            else
            {
                if (File.Exists(Application.StartupPath + @"\FilePathCheckerModuleExample.dll"))
                {
                    RegistryKey regKey = Registry.CurrentUser.CreateSubKey(@"Software\HNC\HwpCtrl\Modules", RegistryKeyPermissionCheck.ReadWriteSubTree);
                    try {
                        regKey.SetValue("FilePathCheckerModuleExample", Application.StartupPath + @"\FilePathCheckerModuleExample.dll", RegistryValueKind.String);
                    }
                    catch (Exception e1)
                    {
                        MessageBox.Show(e1.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    finally
                    {
                        regKey.Close();
                    }
                }

                pictureBox1.Image = DocConvert_Util.Properties.Resources.switch_on;
                HWPREGDLL = true;
            }
        }
    }
}
