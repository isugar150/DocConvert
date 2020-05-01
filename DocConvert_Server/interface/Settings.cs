using DocConvert_Core.IniLib;

namespace DocConvert_Core.interfaces
{
    interface iniproperties
    {
        string LicenseKEY { get; set; }
        string ServerName { get; set; }
        string BindIP { get; set; }
        int WebSocketPort { get; set; }
        int FileServerPort { get; set; }
        int DisplayLogCnt { get; set; }
        string ClientKEY { get; set; }
        string DataPath { get; set; }
        bool OfficeDebugModeYn { get; set; }
        bool FollowTailYn { get; set; }
        bool CleanWorkspaceSchedulerYn { get; set; }
        int CleanWorkspaceDay { get; set; }
        bool CleanLogSchedulerYn { get; set; }
        int CleanLogDay { get; set; }
        bool ChromiumCaptureYn { get; set; }
        int WebCaptureTimeout { get; set; }
    }

    public class iniProperties : iniproperties
    {
        private string _LicenseKEY;
        private string _ServerName;
        private string _BindIP;
        private int _WebSocketPort;
        private int _FileServerPort;
        private int _DisplayLogCnt;
        private string _ClientKEY;
        private string _DataPath;
        private bool _OfficeDebugModeYn;
        private bool _FollowTailYn;
        private bool _CleanWorkspaceSchedulerYn;
        private int _CleanWorkspaceDay;
        private bool _CleanLogSchedulerYn;
        private int _CleanLogDay;
        private bool _ChromiumCaptureYn { get; set; }
        private int _WebCaptureTimeout;
        public string LicenseKEY { get { return _LicenseKEY; } set { _LicenseKEY = value; } }
        public string ServerName { get { return _ServerName; } set { _ServerName = value; } }
        public string BindIP { get { return _BindIP; } set { _BindIP = value; } }
        public int WebSocketPort { get { return _WebSocketPort; } set { _WebSocketPort = value; } }
        public int FileServerPort { get { return _FileServerPort; } set { _FileServerPort = value; } }
        public int DisplayLogCnt { get { return _DisplayLogCnt; } set { _DisplayLogCnt = value; } }
        public string ClientKEY { get { return _ClientKEY; } set { _ClientKEY = value; } }
        public string DataPath { get { return _DataPath; } set { _DataPath = value; } }
        public bool OfficeDebugModeYn { get { return _OfficeDebugModeYn; } set { _OfficeDebugModeYn = value; } }
        public bool FollowTailYn { get { return _FollowTailYn; } set { _FollowTailYn = value; } }
        public bool CleanWorkspaceSchedulerYn { get { return _CleanWorkspaceSchedulerYn; } set { _CleanWorkspaceSchedulerYn = value; } }
        public int CleanWorkspaceDay { get { return _CleanWorkspaceDay; } set { _CleanWorkspaceDay = value; } }
        public bool CleanLogSchedulerYn { get { return _CleanLogSchedulerYn; } set { _CleanLogSchedulerYn = value; } }
        public int CleanLogDay { get { return _CleanLogDay; } set { _CleanLogDay = value; } }
        public bool ChromiumCaptureYn { get { return _ChromiumCaptureYn; } set { _ChromiumCaptureYn = value; } }
        public int WebCaptureTimeout { get { return _WebCaptureTimeout; } set { _WebCaptureTimeout = value; } }
    }

    public class Setting
    {
        public static void createSetting()
        {
            IniFile setting = new IniFile();

            setting["DC Server"]["LicenseKEY"] = "";
            setting["DC Server"]["ServerName"] = "Jm's DCServer";
            setting["DC Server"]["BindIP"] = "Any";
            setting["DC Server"]["WebSocketPort"] = "12000";
            setting["DC Server"]["FileServerPort"] = "12100";
            setting["DC Server"]["DisplayLogCnt"] = "300";
            setting["DC Server"]["ClientKEY"] = "B29D00A3 - F825 - 4EB7 - 93C1 - A77F5E31A7C2";
            setting["DC Server"]["DataPath"] = "C:\\Data";
            setting["DC Server"]["OfficeDebugModeYn"] = "Y";
            setting["DC Server"]["FollowTailYn"] = "Y";
            setting["DC Server"]["CleanWorkspaceSchedulerYn"] = "Y";
            setting["DC Server"]["CleanWorkspaceDay"] = "3";
            setting["DC Server"]["CleanLogSchedulerYn"] = "Y";
            setting["DC Server"]["CleanLogDay"] = "10";
            setting["DC Server"]["ChromiumCaptureYn"] = "n";
            setting["DC Server"]["WebCaptureTimeout"] = "30";

            setting.Save("./DocConvert_Server.ini");
        }
    }
}
