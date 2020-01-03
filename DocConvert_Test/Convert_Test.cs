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

namespace DocConvert
{
    public partial class Convert_Test : Form
    {
        IPAddress ipAddress;
        public Convert_Test()
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
            openFileDialog1.Filter = "지원하는 형식 (*.docx;*.doc;*.xlsx;*.xls;*.ppt;*.pptx)|*.docx;*.doc;*.xlsx;*.xls;*.ppt;*.pptx|All Files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;

            if(openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
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
    }
}
