import org.apache.commons.net.PrintCommandListener;
import org.apache.commons.net.ftp.FTP;
import org.apache.commons.net.ftp.FTPClient;
import org.apache.commons.net.ftp.FTPFile;
import org.apache.commons.net.ftp.FTPReply;

import java.io.*;
import java.util.List;

// 리눅스에서 사용할때 vsftpd 설치해야함.
public class FTPManager {
    private FTPClient ftpClient = new FTPClient();

    public void FTPManager() {

    }

    /**
     * FTP 서버에 커넥션을 연결합니다.
     *
     * @param serverIP    FTP연결할 서버의 IP
     * @param serverPORT  FTP연결할 서버의 포트
     * @param ftpUser     FTP연결할 서버의 사용자ID
     * @param ftpPassword FTP연결할 서버의 사용자PW
     * @throws IOException
     */
    public void Connect(String serverIP, int serverPORT, String ftpUser, String ftpPassword) throws IOException {
        ftpClient.addProtocolCommandListener(new PrintCommandListener(new PrintWriter(System.out)));
        int reply;
        ftpClient.connect(serverIP, serverPORT);//호스트 연결
        reply = ftpClient.getReplyCode();
        if (!FTPReply.isPositiveCompletion(reply)) {
            ftpClient.disconnect();
            throw new IOException("Exception in connecting to FTP Server");
        }
        ftpClient.login(ftpUser, ftpPassword);//로그인
        ftpClient.setFileType(FTP.BINARY_FILE_TYPE);
        ftpClient.enterLocalPassiveMode();
        System.out.println(serverIP + ":" + serverPORT + " 해당 FTP서버에 접속하였습니다.");
    }

    /**
     * FTP 서버에 파일을 업로드합니다.
     *
     * @param localFilePath 내 PC에서 업로드할 파일경로(파일명 빼고)
     * @param fileName      내 PC에서 업로드할 파일의 이름
     * @param remoteDir     FTP서버에 업로드할 경로
     * @throws Exception
     */
    public void uploadFile(String localFilePath, String fileName, String remoteDir) throws Exception {
        try (InputStream input = new FileInputStream(new File(localFilePath + File.separator + fileName))) {
            ftpClient.storeFile(remoteDir + fileName, input);
            System.out.println(remoteDir + fileName + " 해당 파일을 업로드 하였습니다.");
        }
    }

    public void downloadFile(String remoteFile, String localFile) throws IOException {
        ftpClient.setFileType(FTP.BINARY_FILE_TYPE);
        ftpClient.setFileTransferMode(FTP.BINARY_FILE_TYPE);

        FileOutputStream fos = new FileOutputStream(localFile);
        ftpClient.retrieveFile(remoteFile, fos);
        System.out.println("파일을 다운로드 하였습니다. " + localFile);
    }

    /**
     * FTP 서버에 연결된 커넥션을 해제합니다.
     *
     * @throws IOException
     */
    public void disConnect() {
        if (ftpClient.isConnected()) {
            try {
                ftpClient.logout();
                ftpClient.disconnect();
                System.out.println("FTP서버에서 접속을 해제하였습니다.");
            } catch (IOException f) {
                f.printStackTrace();
            }
        }
    }
}