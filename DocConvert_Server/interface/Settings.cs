using DocConvert_Core.IniLib;

namespace DocConvert_Core.interfaces
{
    internal interface iniproperties
    {
        string LicenseKEY { get; set; }
        string BindIP { get; set; }
        int WebSocketPort { get; set; }
        int FileServerPort { get; set; }
        int DisplayLogCnt { get; set; }
        string ClientKEY { get; set; }
        string DataPath { get; set; }
        bool OfficeDebugModeYn { get; set; }
        bool FollowTailYn { get; set; }
        string ResponseTimeout { get; set; }
        string SchedulerTime { get; set; }
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
        private string _BindIP;
        private int _WebSocketPort;
        private int _FileServerPort;
        private int _DisplayLogCnt;
        private string _ClientKEY;
        private string _DataPath;
        private bool _OfficeDebugModeYn;
        private bool _FollowTailYn;
        private string _ResponseTimeout;
        private string _SchedulerTime;
        private bool _CleanWorkspaceSchedulerYn;
        private int _CleanWorkspaceDay;
        private bool _CleanLogSchedulerYn;
        private int _CleanLogDay;
        private bool _ChromiumCaptureYn { get; set; }
        private int _WebCaptureTimeout;
        public string LicenseKEY { get { return _LicenseKEY; } set { _LicenseKEY = value; } }
        public string BindIP { get { return _BindIP; } set { _BindIP = value; } }
        public int WebSocketPort { get { return _WebSocketPort; } set { _WebSocketPort = value; } }
        public int FileServerPort { get { return _FileServerPort; } set { _FileServerPort = value; } }
        public int DisplayLogCnt { get { return _DisplayLogCnt; } set { _DisplayLogCnt = value; } }
        public string ClientKEY { get { return _ClientKEY; } set { _ClientKEY = value; } }
        public string DataPath { get { return _DataPath; } set { _DataPath = value; } }
        public bool OfficeDebugModeYn { get { return _OfficeDebugModeYn; } set { _OfficeDebugModeYn = value; } }
        public bool FollowTailYn { get { return _FollowTailYn; } set { _FollowTailYn = value; } }
        public string ResponseTimeout { get { return _ResponseTimeout; } set { _ResponseTimeout = value; } }
        public string SchedulerTime { get { return _SchedulerTime; } set { _SchedulerTime = value; } }
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

            setting["DC Server"]["LicenseKEY"] = ";정품 확인키";
            setting["DC Server"]["BindIP"] = "Any;바인드할 IP주소 (Any: 아무 주소)";
            setting["DC Server"]["WebSocketPort"] = "12000;WebSocket Listning할 포트";
            setting["DC Server"]["FileServerPort"] = "12100;감시할 파일전송 포트";
            setting["DC Server"]["DisplayLogCnt"] = "300;리스트박스에 표시할 최대 로그카운트 (단위: 개)";
            setting["DC Server"]["ClientKEY"] = "B29D00A3 - F825 - 4EB7 - 93C1 - A77F5E31A7C2;해당 클라이언트가 허용한 클라이언트인지 확인하는 문자열, 클라이언트와 동일하게 맞추면됨.";
            setting["DC Server"]["DataPath"] = "C:\\Data;작업폴더 실제 경로";
            setting["DC Server"]["OfficeDebugModeYn"] = "Y;오피스 변환시 변환하는 화면보이기 (Y:사용) (n:사용안함)";
            setting["DC Server"]["FollowTailYn"] = "Y;프로그램 시작시 FollowTail 체크 (Y:사용) (n:사용안함)";
            setting["DC Server"]["ResponseTimeout"] = "2,0;서버에서 클라이언트로 응답 Timeout 서버 환경에따라 설정 (분,초)";
            setting["DC Server"]["SchedulerTime"] = "1,0;스케줄러 동작시간(매일) 24시간중 (시간,분)";
            setting["DC Server"]["CleanWorkspaceSchedulerYn"] = "Y;작업공간 정리 스케줄러 사용 (Y:사용) (n:사용안함)";
            setting["DC Server"]["CleanWorkspaceDay"] = "3;작업공간 정리시 설정한 오래된 일수가 지난 파일 삭제 (단위: 일)";
            setting["DC Server"]["CleanLogSchedulerYn"] = "Y;로그 정리 스케줄러 사용 (Y:사용) (n:사용안함)";
            setting["DC Server"]["CleanLogDay"] = "10;로그 정리시 설정한 오래된 일수가 지난 파일 삭제 (단위: 일)";
            setting["DC Server"]["ChromiumCaptureYn"] = "n;(Y:크로미움기반 웹 문서 캡쳐 사용) (n:Phantom JS를 이용한 웹 문서 캡쳐)";
            setting["DC Server"]["WebCaptureTimeout"] = "30;특정 시간이지나면 캡쳐중인 콘솔창 자동 Kill (단위: 초)";

            setting.Save("./DocConvert_Server.ini");
        }
    }
}
