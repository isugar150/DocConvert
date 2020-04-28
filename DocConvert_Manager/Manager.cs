using NLog;
using NLog.Internal;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DocConvert_Manager
{
    public partial class Manager : Form
    {
        private static Logger logger = LogManager.GetLogger("DocConvert_Manager_Log");
        public Manager()
        {
            InitializeComponent();
            CheckForIllegalCrossThreadCalls = false;
        }

        private void Manager_Load(object sender, EventArgs e)
        {
            /*Configuration configuration = DocConvert_Server.*/
            new Thread(delegate ()
            {
                while (!this.IsDisposed)
                {
                    if (checkBox1.Checked)
                    {

                    }
                }
            }).Start();
        }
    }
}
