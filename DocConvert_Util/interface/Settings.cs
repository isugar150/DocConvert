using DocConvert_Core.IniLib;

namespace DocConvert_Core.interfaces
{
    internal interface iniproperties
    {
        string TargetIP { get; set; }
        int socketPort { get; set; }
        string clientKEY { get; set; }
        bool appvisible { get; set; }
        bool runafter { get; set; }
        bool pagingnum { get; set; }
    }

    public class iniProperties : iniproperties
    {
        private string _TargetIP;
        private int _socketPort;
        private string _clientKEY;
        private bool _appvisible;
        private bool _runafter;
        private bool _pagingnum;
        public string TargetIP { get { return _TargetIP; } set { _TargetIP = value; } }
        public int socketPort { get { return _socketPort; } set { _socketPort = value; } }
        public string clientKEY { get { return _clientKEY; } set { _clientKEY = value; } }
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
            setting["DC Util"]["socketPort"] = "12000";
            setting["DC Util"]["clientKEY"] = "B29D00A3 - F825 - 4EB7 - 93C1 - A77F5E31A7C2";
            setting["DC Util"]["appvisible"] = "N";
            setting["DC Util"]["runafter"] = "N";
            setting["DC Util"]["pagingnum"] = "N";
            setting.Save("./DocConvert_Util.ini");
        }

        public static void updateSetting(iniProperties properties)
        {
            IniFile setting = new IniFile();
            setting["DC Util"]["targetIP"] = properties.TargetIP;
            setting["DC Util"]["socketPort"] = properties.socketPort;
            setting["DC Util"]["clientKEY"] = properties.clientKEY;
            setting["DC Util"]["appvisible"] = properties.appvisible ? "Y" : "n";
            setting["DC Util"]["runafter"] = properties.runafter ? "Y" : "n";
            setting["DC Util"]["pagingnum"] = properties.pagingnum ? "Y" : "n";
            setting.Save("./DocConvert_Util.ini");
        }
    }
}
