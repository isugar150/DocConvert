using NLog;
using System;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace DocConvert_Server
{
    public partial class Form1 : Form
    {
        private static Logger logger = LogManager.GetLogger("DocConvert_Server_Log");
        public Form1()
        {
            InitializeComponent();
            CheckForIllegalCrossThreadCalls = false;
        }

        #region 통신관련

        #endregion

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }

        private void consoleAppend(string text)
        {
            textBox1.AppendText(System.DateTime.Now.ToString("yyyy/MM/dd hh:mm:ss.fff") + "   " + text + "\r\n");
        }
    }
}
