using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocConvert_Server
{
    public class getSettings
    {
        public static int getLongMaxCount
        {
            get
            {
                return Properties.Settings.Default.LogMaxCount;
            }
        }

        public static int getFileServerPORT
        {
            get
            {
                return Properties.Settings.Default.fileServerPORT;
            }
        }

        public static int getSocketPORT
        {
            get
            {
                return Properties.Settings.Default.socketPORT;
            }
        }

        public static string getServerIP
        {
            get
            {
                return Properties.Settings.Default.serverIP;
            }
        }

        public static string getDataPath
        {
            get
            {
                return Properties.Settings.Default.DataPath;
            }
        }
    }
}
