package docconvert;

import me.saro.commons.ftp.FTP;
import org.java_websocket.client.WebSocketClient;
import org.java_websocket.handshake.ServerHandshake;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
import org.json.simple.parser.ParseException;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URI;

public class DocConvert{
    private String host = "127.0.0.1"; // 호스트
    private int port = 12100; // 포트 : 기본포트(FTP 21, FTPS 990, SFTP 22)
    private String user = "user1"; // 유저이름
    private String pass = "1234"; // 암호
    private JSONObject responseData = null;
    public DocConvert(String filePath, String outPath, String fileName, int toImg) throws Exception {
        //초기 데이터 입력
        String sourceFileExten = fileName.substring(fileName.lastIndexOf( "." ), fileName.length());
        String downloadIMGDir = fileName.replace(sourceFileExten, "");
        String downloadPDFName = fileName.replace(sourceFileExten, ".pdf");

        // 기본 디렉토리 생성
        new File(filePath).mkdirs();

        // 기본 전송 파일 생성
        try (FileOutputStream fos = new FileOutputStream(filePath + fileName)) {
            fos.write("the test file".getBytes());
        }

        // FTP.openFTP : FTP
        // FTP.openFTPS : FTPS
        // FTP.openSFTP : SFTP
        try (FTP ftp = FTP.openFTP(host, port, user, pass)) {

            System.out.println("==================================");
            System.out.println("## 현재위치");
            System.out.println(ftp.path());
            System.out.println("## 디렉토리 목록");
            ftp.listDirectories().forEach(e -> System.out.println(e));
            System.out.println("## 파일 목록");
            ftp.listFiles().forEach(e -> System.out.println(e));
            System.out.println("==================================");

            // tmp 경로로 이동
            ftp.path("/tmp/");

            // 파일전송
            ftp.send(fileName, new File(filePath + fileName));
            /*ftp.send(fileName, new File(filePath + fileName));*/

            // 웹소켓 기능
            // 서버 전송전 데이터
            JSONObject requestMsg = new JSONObject();
            requestMsg.put("KEY", "B29D00A3 - F825 - 4EB7 - 93C1 - A77F5E31A7C2");
            requestMsg.put("FileName", fileName);
            requestMsg.put("ConvertIMG", toImg);
            requestMsg.put("DocPassword", "");

            // 웹 소켓으로 서버에게 데이터 전송
            WebSocketClient webSocketClient = new WebSocketClient(new URI("ws://10.1.3.167:12005")) {

                @Override
                public void onOpen(ServerHandshake serverHandshake) {
                    System.out.println("[Client ==> Server]\r\n" + requestMsg.toJSONString());
                    this.send(requestMsg.toJSONString());
                }

                @Override
                public void onMessage(String message) {
                    this.close();
                    try {
                        responseData = (JSONObject)new JSONParser().parse(message);
                    } catch (ParseException e) {
                        e.printStackTrace();
                    }

                    System.out.println("[DocConvert] 소켓 연결됨.");

                    System.out.println("[Server ==> Client]\r\n" + responseData.toJSONString());

                    // 다운로드
                    try {
                        ftp.path(responseData.get("URL").toString());
                        ftp.recv(downloadPDFName, new File(filePath + downloadPDFName));
                        ftp.path(responseData.get("URL").toString() + "/" + downloadIMGDir);
                        ftp.recv("tmp", new File(filePath + fileName)); // is not file, return false; not recv
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                    ftp.close();
                }

                @Override
                public void onClose(int code, String reason, boolean remote) {
                    System.out.println(String.format("[DocConvert] code: %s, reason: %s, remote: %s", code, reason, remote));
                    // 서버 연결 종료 후 동작 정의
                }

                @Override
                public void onError(Exception ex) {
                    // 예외 발생시 동작 정의
                    System.out.println("[DocConvert][Error] " + ex.getLocalizedMessage());
                }
            };
            // 앞서 정의한 WebSocket 서버에 연결한다.
            webSocketClient.connect();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
