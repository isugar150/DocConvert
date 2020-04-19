package DocConvert_API;

import org.apache.commons.net.PrintCommandListener;
import org.apache.commons.net.ftp.*;

import java.io.*;

// 리눅스에서 사용할때 vsftpd 설치해야함.
public class FTPManager {
    private FTPClient ftpClient = new FTPClient();
    private FTPSClient ftpsClient = new FTPSClient("TLSv1.2", true);

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
        ftpClient.setControlEncoding("UTF-8");
        /*ftpClient.addProtocolCommandListener(new PrintCommandListener(new PrintWriter(System.out)));*/
        int reply;
        System.out.println("FTP서버에 연결을 시도합니다..");
        ftpClient.connect(serverIP, serverPORT);//호스트 연결
        reply = ftpClient.getReplyCode();
        if (!FTPReply.isPositiveCompletion(reply)) {
            ftpClient.disconnect();
            throw new IOException("Exception in connecting to FTP Server");
        }
        ftpClient.login(ftpUser, ftpPassword);//로그인
        ftpClient.setFileType(FTP.BINARY_FILE_TYPE);
        ftpClient.setFileTransferMode(FTP.BINARY_FILE_TYPE);
        ftpClient.enterLocalPassiveMode();
        ftpClient.setControlKeepAliveTimeout((1000*60)*5);
        System.out.println(serverIP + ":" + serverPORT + " 해당 FTP서버에 접속하였습니다.");
    }

    public void ConnectFTPS(String serverIP, int serverPORT, String ftpUser, String ftpPassword) throws IOException{
        System.setProperty("https.protocols", "TLSv1,TLSv1.1,TLSv1.2");
        ftpsClient.setControlEncoding("UTF-8");
        /*ftpsClient.addProtocolCommandListener(new PrintCommandListener(new PrintWriter(System.out)));*/
        int reply;
        System.out.println("FTPS서버에 연결을 시도합니다..");
        ftpsClient.connect(serverIP, serverPORT);//호스트 연결
        reply = ftpsClient.getReplyCode();
        if (!FTPReply.isPositiveCompletion(reply)) {
            ftpsClient.disconnect();
            throw new IOException("Exception in connecting to FTPS Server");
        }
        ftpsClient.login(ftpUser, ftpPassword);//로그인
        ftpsClient.setFileType(FTP.BINARY_FILE_TYPE);
        ftpsClient.setFileTransferMode(FTP.BINARY_FILE_TYPE);
        ftpsClient.enterLocalPassiveMode();
        ftpsClient.setControlKeepAliveTimeout((1000*60)*5);
        ftpsClient.execPBSZ(0);
        ftpsClient.execPROT("P");
        System.out.println(serverIP + ":" + serverPORT + " 해당 FTPS서버에 접속하였습니다.");
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
        } catch(Exception e){
            throw e;
        }
    }

    public void uploadFileFTPS(String localFilePath, String fileName, String remoteDir) throws Exception {
        try (InputStream input = new FileInputStream(new File(localFilePath + File.separator + fileName))) {
            ftpsClient.storeFile(remoteDir + fileName, input);
            System.out.println("[FTPS]" + remoteDir + fileName + " 해당 파일을 업로드 하였습니다.");
        } catch(Exception e){
            throw e;
        }
    }

    /**
     * FTP 서버로부터 파일을 다운로드합니다
     * @param remoteFile 원격지 서버로 부터의 파일 경로
     * @param localFile  다운로드 받을 파일의 경로
     * @throws IOException
     */
    public void downloadFile(String remoteFile, String localFile) throws IOException {
        FileOutputStream fos = new FileOutputStream(localFile);
        ftpClient.retrieveFile(remoteFile, fos);
        System.out.println("파일을 다운로드 하였습니다. " + localFile);
    }

    public void downloadFileFTPS(String remoteFile, String localFile) throws IOException {
        FileOutputStream fos = new FileOutputStream(localFile);
        ftpsClient.retrieveFile(remoteFile, fos);
        System.out.println("파일을 다운로드 하였습니다. " + localFile);
    }

    /**
     * FTP 서버에 업로드한 파일을 삭제합니다
     * @param remoteFile 원격지 서버로 부터의 파일 경로
     * @throws IOException
     */
    public void deleteFile(String remoteFile) throws IOException {
        ftpClient.setFileType(FTP.BINARY_FILE_TYPE);
        ftpClient.setFileTransferMode(FTP.BINARY_FILE_TYPE);
        ftpClient.deleteFile(remoteFile);
    }

    public void deleteFileFTPS(String remoteFile) throws IOException {
        ftpsClient.setFileType(FTP.BINARY_FILE_TYPE);
        ftpsClient.setFileTransferMode(FTP.BINARY_FILE_TYPE);
        ftpsClient.deleteFile(remoteFile);
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

    public void disConnectFTPS() {
        if (ftpsClient.isConnected()) {
            try {
                ftpsClient.logout();
                ftpsClient.disconnect();
                System.out.println("FTPS서버에서 접속을 해제하였습니다.");
            } catch (IOException f) {
                f.printStackTrace();
            }
        }
    }
}