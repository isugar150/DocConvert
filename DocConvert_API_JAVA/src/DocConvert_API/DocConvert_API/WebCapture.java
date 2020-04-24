package DocConvert_API;

import org.java_websocket.client.WebSocketClient;
import org.java_websocket.handshake.ServerHandshake;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
import org.json.simple.parser.ParseException;

import java.io.File;
import java.io.IOException;
import java.net.URI;

public class WebCapture {
    private JSONObject responseData = null;
    private String returnValue = null;
    public String WebCapture_Start(final String Url, final String outPath) throws Exception {
        // 환경설정 읽기
        final getProperties properties = new getProperties();
        properties.readProperties();

        String host = properties.getServerIP(); // 호스트
        int webSocketPort = properties.getServerPORT();
        int ftpPort = properties.getFtpPORT();
        String ftpUser = properties.getFtpUSER(); // 유저이름
        String ftpPass = properties.getFtpPASS(); // 암호
        final boolean isFTPS = properties.getIsFTPS().equals("Y") ? true : false; // FTPS 사용여부
        String clientKEY = properties.getClientKEY();

        final long start = System.currentTimeMillis(); //코드 실행 전에 시간 받아오기

        // FTP변수 초기화
        final FTPManager ftpManager = new FTPManager();
        if(isFTPS)
            ftpManager.ConnectFTPS(host, ftpPort, ftpUser, ftpPass);
        else
            ftpManager.Connect(host, ftpPort, ftpUser, ftpPass);

        // 웹소켓 기능
        // 서버 전송전 데이터
        final JSONObject requestMsg = new JSONObject();
        requestMsg.put("KEY", "B29D00A3 - F825 - 4EB7 - 93C1 - A77F5E31A7C2");
        requestMsg.put("Method", "WebCapture");
        requestMsg.put("URL", Url);

        // 웹 소켓으로 서버에게 데이터 전송
        System.out.println("ws://" + host + ":" + webSocketPort + "에 연결을 시도합니다.");
        WebSocketClient webSocketClient = new WebSocketClient(new URI("ws://" + host + ":" + webSocketPort)) {

            @Override
            public void onOpen(ServerHandshake serverHandshake) {
                System.out.println("[Client ==> Server]\r\n" + requestMsg.toJSONString());
                this.send(requestMsg.toJSONString());
            }

            @Override
            public void onMessage(String message) {
                try {
                    responseData = (JSONObject) new JSONParser().parse(message);

                    System.out.println("[DocConvert] 소켓 연결됨.");

                    System.out.println("[Server ==> Client]\r\n" + responseData.toJSONString());

                    String downloadUrl = responseData.get("URL").toString();
                    String localUrl = outPath + File.separator + new File(downloadUrl).getName();

                    System.out.println("downloadUrl: " + downloadUrl);
                    // PDF 다운로드
                    if(isFTPS)
                        ftpManager.downloadFileFTPS(responseData.get("URL").toString().replace("\\", "/"), outPath + File.separator + new File(outPath).getName());
                    else
                        ftpManager.downloadFile(responseData.get("URL").toString().replace("\\", "/"), outPath + File.separator + new File(outPath).getName());
                } catch (ParseException | IOException e) {
                    e.printStackTrace();
                }
                if(isFTPS)
                    ftpManager.disConnectFTPS();
                else
                    ftpManager.disConnect();
                returnValue = responseData.toJSONString();
                long end = System.currentTimeMillis(); //프로그램이 끝나는 시점 계산
                System.out.println( "실행 시간 : " + ( end - start )/1000.0 +"초"); //실행 시간 계산 및 출력
            }

            @Override
            public void onClose(int code, String reason, boolean remote) {
                // 서버 연결 종료 후 동작 정의
                System.out.println(String.format("[DocConvert] code: %s, reason: %s, remote: %s", code, reason, remote));
                if(isFTPS)
                    ftpManager.disConnectFTPS();
                else
                    ftpManager.disConnect();
            }

            @Override
            public void onError(Exception ex) {
                // 예외 발생시 동작 정의
                if(isFTPS)
                    ftpManager.disConnectFTPS();
                else
                    ftpManager.disConnect();
                ex.printStackTrace();
            }
        };

        webSocketClient.connect();
        while(!webSocketClient.isClosed())
            Thread.sleep(300);
        return returnValue;
    }
}
