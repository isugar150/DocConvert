using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Newtonsoft.Json.Linq;
using NLog;

namespace DocConvert
{
    public partial class Converter_Server : Form
    {
        private static Logger logger = LogManager.GetLogger("DocConvert_Server_Log");
        JObject Setting = null;
        public Converter_Server()
        {
            InitializeComponent();
            CheckForIllegalCrossThreadCalls = false;
        }

        [Obsolete]
        private void Converter_Server_Load(object sender, EventArgs e)
        {
            if(File.Exists(Application.StartupPath + @"\Settings.json")){
                Setting = JObject.Parse(File.ReadAllText(Application.StartupPath + @"\Settings.json"));
                textBox1.AppendText("설정 파일을 불러왔습니다.\r\n");
                textBox1.AppendText("내부 IP: " + Dns.GetHostByName(Dns.GetHostName()).AddressList[Dns.GetHostByName(Dns.GetHostName()).AddressList.Length-1].ToString() + "\r\n");
                textBox1.AppendText("설정된 포트번호: " + Setting["port"] + "\r\n");
                textBox1.AppendText("저장 경로: " + Setting["path"] + "\r\n");
            }
            else
            {
                textBox1.AppendText("설정 파일이 존재하지 않습니다. 설정 먼저 해주세요.\r\n");
            }
        }

        private void 설정ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (var SettingsForm = new Settings())
            {
                SettingsForm.Setting = Setting;
                SettingsForm.ShowDialog(this);
                Setting = SettingsForm.Setting;

                textBox1.AppendText("설정된 포트번호: " + Setting["port"] + "\r\n");
                textBox1.AppendText("저장 경로: " + Setting["path"] + "\r\n");
            }
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
