using DocConvert_Core.IniLib;

namespace DocConvert_Core.interfaces
{
    internal interface iniproperties
    {
        string targetPath { get; set; }
        int refreshCycle { get; set; }
        string runOption { get; set; }
        bool autoRestart { get; set; }
        string autoRestartTime { get; set; }
    }

    public class iniProperties : iniproperties
    {
        private string _targetPath;
        private int _refreshCycle;
        private string _runOption;
        private bool _autoRestart;
        private string _autoRestartTime;
        public string targetPath { get { return _targetPath; } set { _targetPath = value; } }
        public int refreshCycle { get { return _refreshCycle; } set { _refreshCycle = value; } }
        public string runOption { get { return _runOption; } set { _runOption = value; } }
        public bool autoRestart { get { return _autoRestart; } set { _autoRestart = value; } }
        public string autoRestartTime { get { return _autoRestartTime; } set { _autoRestartTime = value; } }
    }

    public class Setting
    {
        public static void createSetting()
        {
            IniFile setting = new IniFile();

            setting["DC Manager"]["targetPath"] = "./DocConvert_Server.exe";
            setting["DC Manager"]["refreshCycle"] = "1000";
            setting["DC Manager"]["runOption"] = "";
            setting["DC Manager"]["autoRestart"] = "n";
            setting["DC Manager"]["autoRestartTime"] = "2,0,0";

            setting.Save("./DocConvert_Manager.ini");
        }
    }
}
