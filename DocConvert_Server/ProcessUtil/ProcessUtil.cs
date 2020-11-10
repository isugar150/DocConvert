using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocConvert_Server.ProcessUtil
{
    class ProcessUtil
    {
        public static string RunProcess(string excuteFile, string arguments)
        {
            string result = "";

            try
            {
                Process pro = new Process();
                pro.StartInfo.FileName = excuteFile;
                pro.StartInfo.Arguments = arguments;
                pro.StartInfo.UseShellExecute = false;
                pro.StartInfo.RedirectStandardOutput = true;
                pro.StartInfo.CreateNoWindow = true;
                pro.EnableRaisingEvents = true;

                pro.Start();

                result = pro.StandardOutput.ReadToEnd();
                pro.WaitForExit();

                pro.Close();
            } catch(Exception e)
            {
                result = e.Message;
            }

            return result;
        }
    }
}
