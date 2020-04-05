import org.java_websocket.client.WebSocketClient;
import org.java_websocket.handshake.ServerHandshake;
import org.json.simple.JSONObject;

import java.net.URI;
import java.net.URISyntaxException;
import java.util.ArrayList;

public class Test01 {
    public static void main(String args[]) throws URISyntaxException {
        JSONObject requestMsg = new JSONObject();
        requestMsg.put("KEY", "B29D00A3 - F825 - 4EB7 - 93C1 - A77F5E31A7C2");
        requestMsg.put("FileName", "Welcome to Hwp.hwp");
        requestMsg.put("ConvertIMG", 0);
        requestMsg.put("DocPassword", "");

        // https://websocket.org/echo.html가 제공하는 WebSocket 에코 서버에서 기능 테스트
        WebSocketClient webSocketClient = new WebSocketClient(new URI("ws://127.0.0.1:12005")) {

            @Override
            public void onOpen(ServerHandshake serverHandshake) {
                this.send(requestMsg.toJSONString());
            }

            @Override
            public void onMessage(String message) {
                System.out.println("서버에서 받은 메시지\r\n" + message);
                this.close();
            }

            @Override
            public void onClose(int code, String reason, boolean remote) {
                // 서버 연결 종료 후 동작 정의
            }

            @Override
            public void onError(Exception ex) {
                // 예외 발생시 동작 정의
            }
        };
        // 앞서 정의한 WebSocket 서버에 연결한다.
        webSocketClient.connect();
    }
}
