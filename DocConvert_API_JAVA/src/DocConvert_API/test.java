import org.json.simple.JSONObject;

import java.io.*;
import java.net.Socket;
import java.net.SocketException;

public class test {
    public static void main(String args[]){
        try {
            Socket socket = new Socket("localhost", 12000);
            socket.setSoTimeout(30000);
            OutputStream output = socket.getOutputStream();
            String realStr = "";
            byte[] data = realStr.getBytes(); //getBytes() 메서드를 사용 해 문자열을 Byte로 바꿔준다
            output.write(data);
            PrintWriter writer = new PrintWriter(output, true); //true 인수는 메소드 호출 후에 데이터 자동비우기 설정입니다.


            final JSONObject requestMsg = new JSONObject();
            requestMsg.put("KEY", "");
            requestMsg.put("Method", "DocConvert");
            requestMsg.put("FileName", "test.docx");
            requestMsg.put("DocPassword", null);
            requestMsg.put("DRM_UseYn", "n");
            requestMsg.put("DRM_Type", "");

            writer.println(requestMsg);

            InputStream input = socket.getInputStream();
            input.read(data);

            BufferedReader reader = new BufferedReader(new InputStreamReader(input, "UTF-8"));
            String character = reader.readLine();

            String line;
            try{
                while ((line = reader.readLine()) != null) {
                    System.out.println(line);
                    socket.close();
                }
            } catch (SocketException e1){ }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
