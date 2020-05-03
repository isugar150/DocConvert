using DocConvert_Core.IniLib;

namespace DocConvert_Core.interfaces
{
    internal interface iniproperties
    {
        string targetPath { get; set; }
        bool minimized { get; set; }
        int refreshCycle { get; set; }
        string runOption { get; set; }
    }

    public class iniProperties : iniproperties
    {
        private string _targetPath;
        private bool _minimized;
        private int _refreshCycle;
        private string _runOption;
        public string targetPath { get { return _targetPath; } set { _targetPath = value; } }
        public bool minimized { get { return _minimized; } set { _minimized = value; } }
        public int refreshCycle { get { return _refreshCycle; } set { _refreshCycle = value; } }
        public string runOption { get { return _runOption; } set { _runOption = value; } }
    }

    public class Setting
    {
        public static void createSetting()
        {
            IniFile setting = new IniFile();

            setting["DC Manager"]["targetPath"] = "./DocConvert_Server.exe";
            setting["DC Manager"]["minimized"] = "n";
            setting["DC Manager"]["refreshCycle"] = "1000";
            setting["DC Manager"]["runOption"] = "";

            setting.Save("./DocConvert_Manager.ini");
        }
    }
}
