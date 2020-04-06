package docconvert;

import org.java_websocket.client.WebSocketClient;
import org.java_websocket.handshake.ServerHandshake;
import org.json.simple.JSONObject;

import java.net.URI;
import java.net.URISyntaxException;

public class webSocket {
    public void sendData(JSONObject requestMsg) throws URISyntaxException {
        // 웹 소켓으로 서버에게 데이터 전송
        WebSocketClient webSocketClient = new WebSocketClient(new URI("ws://127.0.0.1:12005")) {

            @Override
            public void onOpen(ServerHandshake serverHandshake) {
                this.send(requestMsg.toJSONString());
            }

            @Override
            public void onMessage(String message) {
                System.out.println("[DocConvert]서버에서 받은 메시지\r\n" + message);
                this.close();
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
        System.out.println("[DocConvert] 소켓 연결됨.");
    }
}