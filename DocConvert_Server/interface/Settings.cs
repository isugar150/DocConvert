using DocConvert_Core.IniLib;

namespace DocConvert_Core.interfaces
{
    internal interface iniproperties
    {
        string BindIP { get; set; }
        int WebSocketPort { get; set; }
        int DisplayLogCnt { get; set; }
        string ClientKEY { get; set; }
        string DataPath { get; set; }

        bool OfficeDebugModeYn { get; set; }
        bool ShowTextBoxYn { get; set; }
        bool FollowTailYn { get; set; }
        string ResponseTimeout { get; set; }

        bool FileServerUseYn { get; set; }
        string Http_Prefix { get; set; }
        int FileServerPort { get; set; }

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
        private string _BindIP;
        private int _WebSocketPort;
        private int _DisplayLogCnt;
        private string _ClientKEY;
        private string _DataPath;

        private bool _OfficeDebugModeYn;
        private bool _ShowTextBoxYn;
        private bool _FollowTailYn;
        private string _ResponseTimeout;

        private bool _FileServerUseYn;
        private string _Http_Prefix;
        private int _FileServerPort;

        private string _SchedulerTime;
        private bool _CleanWorkspaceSchedulerYn;
        private int _CleanWorkspaceDay;
        private bool _CleanLogSchedulerYn;
        private int _CleanLogDay;

        private bool _DRM_useYn;
        private string _DRM_Path;
        private string _DRM_Result;
        private string _DRM_Args;



        public string BindIP { get { return _BindIP; } set { _BindIP = value; } }
        public int WebSocketPort { get { return _WebSocketPort; } set { _WebSocketPort = value; } }
        public int DisplayLogCnt { get { return _DisplayLogCnt; } set { _DisplayLogCnt = value; } }
        public string ClientKEY { get { return _ClientKEY; } set { _ClientKEY = value; } }
        public string DataPath { get { return _DataPath; } set { _DataPath = value; } }

        public bool OfficeDebugModeYn { get { return _OfficeDebugModeYn; } set { _OfficeDebugModeYn = value; } }
        public bool ShowTextBoxYn { get { return _ShowTextBoxYn; } set { _ShowTextBoxYn = value; } }
        public bool FollowTailYn { get { return _FollowTailYn; } set { _FollowTailYn = value; } }
        public string ResponseTimeout { get { return _ResponseTimeout; } set { _ResponseTimeout = value; } }

        public bool FileServerUseYn { get { return _FileServerUseYn; } set { _FileServerUseYn = value; } }
        public string Http_Prefix { get { return _Http_Prefix; } set { _Http_Prefix = value; } }
        public int FileServerPort { get { return _FileServerPort; } set { _FileServerPort = value; } }

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

            setting["Common"]["BindIP"] = "Any    ;바인드할 IP주소 (Any: 아무 주소)";
            setting["Common"]["WebSocketPort"] = "12000    ;WebSocket Listning할 포트";
            setting["Common"]["DisplayLogCnt"] = "300      ;리스트박스에 표시할 최대 로그카운트 (단위: 개)";
            setting["Common"]["ClientKEY"] = "B29D00A3 - F825 - 4EB7 - 93C1 - A77F5E31A7C2     ;해당 클라이언트가 허용한 클라이언트인지 확인하는 문자열, 클라이언트와 동일하게 맞추면됨.";
            setting["Common"]["DataPath"] = "C:\\Data      ;작업폴더 실제 경로";

            setting["Options"]["OfficeDebugModeYn"] = "Y        ;오피스 변환시 변환하는 화면보이기 (Y:사용) (n:사용안함)";
            setting["Options"]["ShowTextBoxYn"] = "Y      ;프로그램 시작시 TextBox 체크 (Y:사용) (n:사용안함)";
            setting["Options"]["FollowTailYn"] = "n        ;프로그램 시작시 FollowTail 체크 (Y:사용) (n:사용안함)";
            setting["Options"]["ResponseTimeout"] = "2,0       ;서버에서 클라이언트로 응답 Timeout 서버 환경에따라 설정 (분,초)";

            setting["WebDAV"]["FileServerUseYn"] = "Y         ;WebDAV 서버 사용 여부";
            setting["WebDAV"]["Http Prefix"] = "http://127.0.0.1:12100/         ;http Prefix (실제 접속 URL)";
            setting["WebDAV"]["FileServerPort"] = "12100         ;WebDAV를 이용한 파일 포트";

            setting["Scheduler"]["SchedulerTime"] = "1,0       ;스케줄러 동작시간(매일) 24시간중 (시간,분)";
            setting["Scheduler"]["CleanWorkspaceSchedulerYn"] = "Y         ;작업공간 정리 스케줄러 사용 (Y:사용) (n:사용안함)";
            setting["Scheduler"]["CleanWorkspaceDay"] = "3         ;작업공간 정리시 설정한 오래된 일수가 지난 파일 삭제 (단위: 일)";
            setting["Scheduler"]["CleanLogSchedulerYn"] = "Y       ;로그 정리 스케줄러 사용 (Y:사용) (n:사용안함)";
            setting["Scheduler"]["CleanLogDay"] = "10         ;로그 정리시 설정한 오래된 일수가 지난 파일 삭제 (단위: 일)";

            setting["DRM Setting"]["DRM useYn"] = "n      ;DRM 사용 여부";
            setting["DRM Setting"]["DRM Path"] = "       ;DRM 경로";
            setting["DRM Setting"]["DRM Result"] = "0       ;DRM 성공 시 Result 코드 (해당 코드가 아니면 실패처리)";
            setting["DRM Setting"]["DRM Args"] = "$Full_Path$,$Out_Full_Path$,$DRM_Type$            ;DRM 아규먼트 ','로 구분 { 풀 경로($Full_Path$), 파일 경로($File_Path$), 파일 명($File_Name$), 내보낼 풀 경로($Out_Full_Path$), 변환 타입($DRM_Type$) }";

            setting.Save("./DocConvert_Server.ini");
        }
    }
}
