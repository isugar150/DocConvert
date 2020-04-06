import docconvert.webSocket;
import org.java_websocket.WebSocket;
import org.java_websocket.client.WebSocketClient;
import org.java_websocket.handshake.ServerHandshake;
import org.json.simple.JSONObject;

import java.io.File;
import java.net.URI;
import java.net.URISyntaxException;
import java.util.ArrayList;

import docconvert.FTPUploader;

public class Test01 {
    public static void main(String args[]) throws Exception {
        try {
            //초기 데이터 입력
            String filePath = "C:\\Users\\JMKIM\\Desktop\\Excel.xlsx";

            //FTP로 파일 업로드
            System.out.println("Start");
            FTPUploader ftpUploader = new FTPUploader("127.0.0.1", 12100, "user1", "1234");
            ftpUploader.uploadFile(filePath, new File(filePath).getName(), "/tmp/");
            ftpUploader.disconnect();
            System.out.println("Done");

            // 서버 전송전 데이터
            JSONObject requestMsg = new JSONObject();
            requestMsg.put("KEY", "B29D00A3 - F825 - 4EB7 - 93C1 - A77F5E31A7C2");
            requestMsg.put("FileName", new File(filePath).getName());
            requestMsg.put("ConvertIMG", 0);
            requestMsg.put("DocPassword", "");

            //웹소켓 전송
            new webSocket().sendData(requestMsg);



        } catch(Exception e){
            System.out.println("[Error] " + e.getMessage());
        }
    }
}
