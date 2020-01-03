using Newtonsoft.Json.Linq;
using NLog;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DocConvert
{
    public partial class Settings : Form
    {
        private static Logger logger = LogManager.GetLogger("DocConvert_Server_Log");
        public JObject Setting { get; set; }
        public Settings()
        {
            InitializeComponent();
        }

        private void Settings_Load(object sender, EventArgs e)
        {
            // 기본값 설정
            try
            {
                textBox1.Text = Setting["port"].ToString();
                textBox2.Text = Setting["path"].ToString();
            }
            catch (NullReferenceException)
            {
                Setting = new JObject();
                textBox1.Text = "12500";
                textBox2.Text = @"C:\TDVData";
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (int.Parse(textBox1.Text) >= 65535)
            {
                MessageBox.Show("포트 최대 입력값이 초과하였습니다.\r\n최대 값: 65535",  "경고", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            Setting["port"] = textBox1.Text;
            Setting["path"] = textBox2.Text;
            try
            {
                File.WriteAllText(Application.StartupPath + @"\Settings.json", Setting.ToString());
            }
            catch(Exception e1)
            {
                MessageBox.Show("변환중 오류발생 자세한 내용은 오류로그 참고", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                logger.Info("변환중 오류발생 자세한 내용은 오류로그 참고");
                logger.Error("==================== Method: " + MethodBase.GetCurrentMethod().Name + " ====================");
                logger.Error(new StackTrace(e1, true));
                logger.Error("오류: " + e1.Message);
                logger.Error("==================== End ====================");
                return;
            }
            MessageBox.Show("저장하였습니다.", "정보", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.Dispose();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            using (var dialog = folderBrowserDialog1)
            {
                if(dialog.ShowDialog() == DialogResult.OK)
                {
                    textBox2.Text = dialog.SelectedPath;
                }
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(char.IsDigit(e.KeyChar) || e.KeyChar == Convert.ToChar(Keys.Back)))    //숫자와 백스페이스를 제외한 나머지를 바로 처리
            {
                e.Handled = true;
            }
        }
    }
}
