package DocConvert_API;

import org.java_websocket.client.WebSocketClient;
import org.java_websocket.handshake.ServerHandshake;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
import org.json.simple.parser.ParseException;

import java.io.File;
import java.io.IOException;
import java.net.URI;

import javax.net.*;

public class DocConvert {
    private JSONObject responseData = null;
    private String returnValue = null;

    /**
     * 각종 문서를 PDF 또는 이미지로 변환합니다.
     *
     * @param filePath  파일의 경로(파일명 빼고)
     * @param outPath   내보낼 경로
     * @param fileName  문서파일의 이름
     * @param toImg     이미지 변환 (0:안함) (1:JPG) (2:PNG) (3:BMP)
     * @throws Exception
     */
    public String DocConvert_Start(final String filePath, String outPath, String fileName, final int toImg) throws Exception {
        // 파라미터 유효성 확인
        if(!new File(filePath).exists())
            throw new IOException();

        // 환경설정 읽기
        getProperties properties = new getProperties();
        properties.readProperties();

        String host = properties.getServerIP(); // 호스트
        int webSocketPort = properties.getServerPORT();
        int ftpPort = properties.getFtpPORT();
        String ftpUser = properties.getFtpUSER(); // 유저이름
        String ftpPass = properties.getFtpPASS(); // 암호

        // 변수 데이터 초기화
        String sourceFileExten = fileName.substring(fileName.lastIndexOf("."), fileName.length());
        final String downloadIMGDir = fileName.replace(sourceFileExten, "");
        final String downloadPDFName = fileName.replace(sourceFileExten, ".pdf");

        // FTP변수 초기화
        final FTPManager ftpManager = new FTPManager();
        ftpManager.Connect(host, ftpPort, ftpUser, ftpPass);

        // 변환할 파일 업 로드
        try{
            ftpManager.uploadFile(filePath, fileName, "/tmp/");
        } catch(Exception e){
            e.printStackTrace();
            return returnValue;
        }

        // 웹소켓 기능
        // 서버 전송전 데이터
        final JSONObject requestMsg = new JSONObject();
        requestMsg.put("KEY", "B29D00A3 - F825 - 4EB7 - 93C1 - A77F5E31A7C2");
        requestMsg.put("FileName", fileName);
        requestMsg.put("ConvertIMG", toImg);
        requestMsg.put("DocPassword", "");

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

                    System.out.println("remoteFile: " + responseData.get("URL").toString() + "/" + downloadPDFName);
                    System.out.println("localFile: " + filePath + File.separator + downloadPDFName);
                    // PDF 다운로드
                    ftpManager.downloadFile(responseData.get("URL").toString() + "/" + downloadPDFName, filePath + File.separator + downloadPDFName);
                    if(toImg != 0){
                        String imgExtension = null;
                        if(Integer.parseInt(requestMsg.get("ConvertIMG").toString()) == 1)
                            imgExtension = ".jpg";
                        else if(Integer.parseInt(requestMsg.get("ConvertIMG").toString()) == 2)
                            imgExtension = ".png";
                        else if(Integer.parseInt(requestMsg.get("ConvertIMG").toString()) == 3)
                            imgExtension = ".bmp";

                        new File(filePath + File.separator + downloadIMGDir).mkdirs();
                        for(int i = 0; i < Integer.parseInt(responseData.get("convertImgCnt").toString()); i++){
                            ftpManager.downloadFile(responseData.get("URL").toString() + "/" + downloadIMGDir + "/" + (i + 1) + imgExtension, filePath + File.separator + File.separator + downloadIMGDir + File.separator + (i + 1) + imgExtension);
                        }
                    }
                } catch (ParseException | IOException e) {
                    e.printStackTrace();
                }
                ftpManager.disConnect();
                returnValue = responseData.toJSONString();
            }

            @Override
            public void onClose(int code, String reason, boolean remote) {
                // 서버 연결 종료 후 동작 정의
                System.out.println(String.format("[DocConvert] code: %s, reason: %s, remote: %s", code, reason, remote));
                ftpManager.disConnect();
            }

            @Override
            public void onError(Exception ex) {
                // 예외 발생시 동작 정의
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
