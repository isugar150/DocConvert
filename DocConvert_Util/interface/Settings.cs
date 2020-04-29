using DocConvert_Core.IniLib;

namespace DocConvert_Core.interfaces
{
    interface iniproperties
    {
        string TargetIP { get; }
        string ftpUser { get; }
        string ftpPwd { get; }
        int socketPort { get; }
        int filePort { get; }
        bool isFTPS { get; }
        bool appvisible { get; }
        bool runafter { get; }
        bool pagingnum { get; }
    }

    public class iniProperties : iniproperties
    {
        private string _TargetIP;
        private string _ftpUser;
        private string _ftpPwd;
        private int _socketPort;
        private int _filePort;
        private bool _isFTPS;
        private bool _appvisible;
        private bool _runafter;
        private bool _pagingnum;
        public string TargetIP { get { return _TargetIP; } set { _TargetIP = value; } }
        public string ftpUser { get { return _ftpUser; } set { _ftpUser = value; } }
        public string ftpPwd { get { return _ftpPwd; } set { _ftpPwd = value; } }
        public int socketPort { get { return _socketPort; } set { _socketPort = value; } }
        public int filePort { get { return _filePort; } set { _filePort = value; } }
        public bool isFTPS { get { return _isFTPS; } set { _isFTPS = value; } }
        public bool appvisible { get { return _appvisible; } set { _appvisible = value; } }
        public bool runafter { get { return _runafter; } set { _runafter = value; } }
        public bool pagingnum { get { return _pagingnum; } set { _pagingnum = value; } }
    }

    public class Setting
    {
        public static void createSetting()
        {
            IniFile setting = new IniFile();

            setting["DC Util"]["targetIP"] = "127.0.0.1";
            setting["DC Util"]["ftpUser"] = "user1";
            setting["DC Util"]["ftpPwd"] = "1234";
            setting["DC Util"]["socketPort"] = "12000";
            setting["DC Util"]["filePort"] = "12100";
            setting["DC Util"]["isFTPS"] = "N";
            setting["DC Util"]["appvisible"] = "N";
            setting["DC Util"]["runafter"] = "N";
            setting["DC Util"]["pagingnum"] = "N";
            setting.Save("./DocConvert_Util.ini");
        }

        public static void updateSetting(iniProperties properties)
        {
            IniFile setting = new IniFile();
            setting["DC Util"]["targetIP"] = properties.TargetIP;
            setting["DC Util"]["ftpUser"] = properties.ftpUser;
            setting["DC Util"]["ftpPwd"] = properties.ftpPwd;
            setting["DC Util"]["socketPort"] = properties.socketPort;
            setting["DC Util"]["filePort"] = properties.filePort;
            setting["DC Util"]["isFTPS"] = properties.isFTPS ? "Y" : "n";
            setting["DC Util"]["appvisible"] = properties.appvisible ? "Y" : "n";
            setting["DC Util"]["runafter"] = properties.runafter ? "Y" : "n";
            setting["DC Util"]["pagingnum"] = properties.pagingnum ? "Y" : "n";
            setting.Save("./DocConvert_Util.ini");
        }
    }
}
