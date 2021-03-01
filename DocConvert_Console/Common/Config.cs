using DocConvert_Core.IniLib;
using System;

namespace DocConvert.Common
{
    internal interface iniproperties
    {
        string Bind_IP { get; set; }
        int Socket_Port { get; set; }
        string Client_KEY { get; set; }
        string Workspace_Directory { get; set; }
        string Product_Name { get; set; }

        string SchedulerTime { get; set; }
        bool CleanWorkspaceSchedulerYn { get; set; }
        int CleanWorkspaceDay { get; set; }
        bool CleanLogSchedulerYn { get; set; }
        int CleanLogDay { get; set; }

        bool DRM_useYn { get; set; }
        string DRM_Path { get; set; }
        string DRM_Result { get; set; }
        string DRM_Args { get; set; }
    }

    public class iniProperties : iniproperties
    {
        private string _Bind_IP = "0.0.0.0";
        private int _Socket_Port = 12000;
        private string _Client_KEY = "";
        private string _Workspace_Directory = "C:\\workspace";
        private string _Product_Name = "DocConvert";

        private string _SchedulerTime = "1,0";
        private bool _CleanWorkspaceSchedulerYn = false;
        private int _CleanWorkspaceDay = 3;
        private bool _CleanLogSchedulerYn = false;
        private int _CleanLogDay = 3;

        private bool _DRM_useYn = false;
        private string _DRM_Path = "";
        private string _DRM_Result = "";
        private string _DRM_Args = "";



        public string Bind_IP { get { return _Bind_IP; } set { _Bind_IP = value; } }
        public int Socket_Port { get { return _Socket_Port; } set { _Socket_Port = value; } }
        public string Client_KEY { get { return _Client_KEY; } set { _Client_KEY = value; } }
        public string Workspace_Directory { get { return _Workspace_Directory; } set { _Workspace_Directory = value; } }
        public string Product_Name { get { return _Product_Name; } set { _Product_Name = value; } }

        public string SchedulerTime { get { return _SchedulerTime; } set { _SchedulerTime = value; } }
        public bool CleanWorkspaceSchedulerYn { get { return _CleanWorkspaceSchedulerYn; } set { _CleanWorkspaceSchedulerYn = value; } }
        public int CleanWorkspaceDay { get { return _CleanWorkspaceDay; } set { _CleanWorkspaceDay = value; } }
        public bool CleanLogSchedulerYn { get { return _CleanLogSchedulerYn; } set { _CleanLogSchedulerYn = value; } }
        public int CleanLogDay { get { return _CleanLogDay; } set { _CleanLogDay = value; } }

        public bool DRM_useYn { get { return _DRM_useYn; } set { _DRM_useYn = value; } }
        public string DRM_Path { get { return _DRM_Path; } set { _DRM_Path = value; } }
        public string DRM_Result { get { return _DRM_Result; } set { _DRM_Result = value; } }
        public string DRM_Args { get { return _DRM_Args; } set { _DRM_Args = value; } }
    }

    public class Setting
    {
        public static void createSetting()
        {
            IniFile setting = new IniFile();

            setting["Common"]["Bind IP"] = "0.0.0.0    ;바인드할 IP주소";
            setting["Common"]["Socket Port"] = "12000    ;HTTP 포트";
            setting["Common"]["Client KEY"] = Guid.NewGuid() + "     ;해당 클라이언트가 허용한 클라이언트인지 확인하는 문자열, 클라이언트와 동일하게 맞추면됨.";
            setting["Common"]["Workspace Directory"] = "C:\\workspace      ;작업폴더 실제 경로";
            setting["Common"]["Product Name"] = "DocConvert      ;콘솔 타이틀";

            setting["Scheduler"]["SchedulerTime"] = "1,0       ;스케줄러 동작시간(매일) 24시간중 (시간,분)";
            setting["Scheduler"]["CleanWorkspaceSchedulerYn"] = "Y         ;작업공간 정리 스케줄러 사용 (Y:사용) (n:사용안함)";
            setting["Scheduler"]["CleanWorkspaceDay"] = "3         ;작업공간 정리시 설정한 오래된 일수가 지난 파일 삭제 (단위: 일)";
            setting["Scheduler"]["CleanLogSchedulerYn"] = "Y       ;로그 정리 스케줄러 사용 (Y:사용) (n:사용안함)";
            setting["Scheduler"]["CleanLogDay"] = "7         ;로그 정리시 설정한 오래된 일수가 지난 파일 삭제 (단위: 일)";

            setting["DRM Setting"]["DRM useYn"] = "n      ;DRM 사용 여부";
            setting["DRM Setting"]["DRM Path"] = "       ;DRM 경로";
            setting["DRM Setting"]["DRM Result"] = "0       ;DRM 성공 시 Result 코드 (해당 코드가 아니면 실패처리)";
            setting["DRM Setting"]["DRM Args"] = "$Full_Path$,$Out_Full_Path$,$DRM_Type$            ;DRM 아규먼트 ','로 구분 { 풀 경로($Full_Path$), 파일 경로($File_Path$), 파일 명($File_Name$), 내보낼 풀 경로($Out_Full_Path$), 변환 타입($DRM_Type$) }";

            setting.Save(Environment.CurrentDirectory + @".\DocConvert.ini");
        }
    }
}
