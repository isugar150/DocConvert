package DocConvert;

import com.sun.corba.se.spi.orbutil.fsm.Input;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;

import java.io.*;
import java.net.Socket;
import java.net.SocketException;

public class DocConvertAPI {
    String IPAddress = "";
    int port = 0;
    int socketTImeout = 60000;
    String clientKey = "";
    public DocConvertAPI(String IPAddress, int port, int socketTImeout, String clientKey){
        this.IPAddress = IPAddress;
        this.port = port;
        this.socketTImeout = socketTImeout;
        this.clientKey = clientKey;
    }

    public String DocConvert(String fileName, int convertImg, String docPassword, String drmUseYn, String drmType) throws Exception{
        Socket socket = new Socket(IPAddress, port);
        socket.setSoTimeout(socketTImeout);
        OutputStream output = socket.getOutputStream();
        String realStr = "";
        byte[] data = realStr.getBytes(); //getBytes() 메서드를 사용 해 문자열을 Byte로 바꿔준다
        output.write(data);
        PrintWriter writer = new PrintWriter(output, true); //true 인수는 메소드 호출 후에 데이터 자동비우기 설정입니다.

        final JSONObject requestMsg = new JSONObject();
        requestMsg.put("ClientKey", clientKey);
        requestMsg.put("Method", "DocConvert");
        requestMsg.put("FileName", fileName);
        requestMsg.put("ConvertImg", convertImg);
        requestMsg.put("DocPassword", docPassword);
        requestMsg.put("DRM_UseYn", drmUseYn);
        requestMsg.put("DRM_Type", drmType);

        writer.println(requestMsg);

        InputStream input = socket.getInputStream();
        input.read(data);

        BufferedReader reader = new BufferedReader(new InputStreamReader(input, "UTF-8"));
        String character = reader.readLine();

        String tmp = "";
        String resultStr = "";
        try{
            while ((tmp = reader.readLine()) != null) {
                resultStr += tmp;
            }
            socket.close();
        } catch (SocketException e1){ }
        return resultStr;
    }

    public static void main(String args[]){
        String clientKey = "c8d959df-226d-4860-bfab-8a75342feb10";
        DocConvertAPI docapi = new DocConvertAPI("127.0.0.1", 12000, 60000, clientKey);
        String fileName = "0_1_A4A99D04196229EDEEFCE02EF2FECEA5.pptx";
        try {
            copyFile("C:\\Users\\JMKIM\\Documents\\Document\\doc\\" + fileName, "C:\\workspace\\tmp\\" + fileName);
            System.out.println(docapi.DocConvert(fileName, 2, "", "n", ""));
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void copyFile(String _srcFile, String _targetFile) throws IOException{
        File srcFile = new File(_srcFile);
        File targetFile = new File(_targetFile);

        FileInputStream in = new FileInputStream(srcFile);
        FileOutputStream out = new FileOutputStream(targetFile);

        int fileByte = 0;
        while((fileByte = in.read()) != -1){
            out.write(fileByte);
        }
        in.close();
        out.close();
    }
}
