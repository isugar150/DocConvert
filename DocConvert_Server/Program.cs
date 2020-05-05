using System;
using System.Threading;
using System.Windows.Forms;

namespace DocConvert_Server
{
    internal static class Program
    {
        /// <summary>
        /// 해당 애플리케이션의 주 진입점입니다.
        /// </summary>
        [STAThread]
        private static void Main(string[] args)
        {
            using (Mutex mutex = new Mutex(false, "Global\\58d05e01-d185-4f89-a413-a59d7e75ec86"))
            {
                if (!mutex.WaitOne(500, false))
                {
                    return;
                }
                /*Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);*/
                Application.Run(new Form1(args));
            }
        }
    }
}
